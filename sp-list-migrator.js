// ======================================================================
// SharePoint List Migrator — Browser Console Version
// Paste this entire script into the browser console while on your
// SharePoint site (source site for export, destination site for import).
// ======================================================================

(function () {
  const enc = encodeURIComponent;
  const headers = { Accept: 'application/json;odata=verbose' };

  function currentSiteUrl() {
    // Best-effort guess of the site URL from the current page.
    const u = new URL(window.location.href);
    const cutMarkers = ['/Lists/', '/SitePages/', '/_layouts/', '/Shared%20Documents/', '/Forms/'];
    let path = u.pathname;
    for (const marker of cutMarkers) {
      const idx = path.indexOf(marker);
      if (idx !== -1) { path = path.substring(0, idx); break; }
    }
    return u.origin + path;
  }

  function downloadJson(obj, filename) {
    const jsonStr = JSON.stringify(obj, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function pickJsonFile() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'application/json';
      input.style.display = 'none';
      document.body.appendChild(input);
      input.addEventListener('change', () => {
        const file = input.files[0];
        input.remove();
        if (!file) { reject(new Error('No file selected.')); return; }
        const reader = new FileReader();
        reader.onload = () => {
          try { resolve(JSON.parse(reader.result)); }
          catch (e) { reject(new Error('Selected file is not valid JSON.')); }
        };
        reader.onerror = () => reject(new Error('Could not read file.'));
        reader.readAsText(file);
      });
      document.body.appendChild(input);
      input.click();
    });
  }

  // --------------------------------------------------------------------
  // GET LISTS
  // --------------------------------------------------------------------
  async function getLists(siteUrl) {
    siteUrl = siteUrl.replace(/\/$/, '');
    const res = await fetch(
      `${siteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Title,ItemCount,BaseTemplate&$orderby=Title`,
      { headers }
    );
    if (!res.ok) throw new Error(`Failed to fetch lists (HTTP ${res.status}). Check the site URL / your permissions.`);
    const data = (await res.json()).d;
    const raw = data.results || data;

    const NOISY_TITLES = new Set([
      'Site Assets', 'Site Pages', 'Style Library', 'Form Templates', 'Master Page Gallery',
      'Theme Gallery', 'Web Part Gallery', 'List Template Gallery', 'Solution Gallery',
      'Composed Looks', 'Reusable Content', 'Content and Structure Reports',
      'Suggested Content Browser Locations', 'Long Running Operation Status',
      'Maintenance Log Library', 'Converted Forms', 'Apps in Testing', 'appdata', 'appfiles',
      'Access Requests', 'Cache Profiles', 'Device Channels', 'Variation Labels'
    ]);

    const TEMPLATE_LABELS = {
      100: 'List', 101: 'Doc Library', 102: 'Survey', 103: 'Links', 104: 'Announcements',
      105: 'Contacts', 106: 'Calendar', 107: 'Tasks', 108: 'Discussion', 109: 'Picture Library',
      115: 'Issue Tracking', 120: 'Custom Grid'
    };

    return raw
      .filter(l => !NOISY_TITLES.has(l.Title))
      .map(l => ({
        Title: l.Title,
        ItemCount: l.ItemCount,
        TypeLabel: TEMPLATE_LABELS[l.BaseTemplate] || 'List'
      }));
  }

  // --------------------------------------------------------------------
  // EXPORT
  // --------------------------------------------------------------------
  async function exportMultipleSchemas(siteUrl, listNames, includeViews) {
    siteUrl = siteUrl.replace(/\/$/, '');

    const webRes = await fetch(`${siteUrl}/_api/web?$select=Id`, { headers });
    const currentWebId = webRes.ok ? (await webRes.json()).d.Id : null;

    const SYSTEM_FIELDS_TO_SKIP = [
      'ContentType', 'Attachments', 'Edit', 'LinkTitleNoMenu', 'LinkTitle',
      'LinkTitle2', 'DocIcon', 'ItemChildCount', 'FolderChildCount',
      'AppAuthor', 'AppEditor', '_UIVersionString', 'InstanceID',
      'Order', 'GUID', 'WorkflowVersion', 'WorkflowInstanceID',
      'ParentVersionString', 'ParentLeafName', 'SelectTitle'
    ];

    const cache = new Map();
    const inProgress = new Set();
    const collected = [];

    async function fetchListWithFields(byId, identifier) {
      const url = byId
        ? `${siteUrl}/_api/web/lists(guid'${identifier}')?$expand=Fields`
        : `${siteUrl}/_api/web/lists/getbytitle('${enc(identifier)}')?$expand=Fields`;
      const res = await fetch(url, { headers });
      if (!res.ok) throw new Error(`Failed to fetch list "${identifier}" (HTTP ${res.status}).`);
      return (await res.json()).d;
    }

    async function exportOneList(byId, identifier) {
      const listData = await fetchListWithFields(byId, identifier);
      const listId = listData.Id;

      if (cache.has(listId)) return cache.get(listId);
      if (inProgress.has(listId)) return { circular: true, listName: listData.Title };
      inProgress.add(listId);

      const rawFields = listData.Fields.results || listData.Fields;
      const fields = [];

      for (const f of rawFields) {
        if (f.Hidden || f.ReadOnlyField || f.Sealed || SYSTEM_FIELDS_TO_SKIP.includes(f.InternalName)) continue;

        const fieldObj = {
          Title: f.Title,
          InternalName: f.InternalName,
          FieldTypeKind: f.FieldTypeKind,
          TypeAsString: f.TypeAsString,
          Required: f.Required,
          SchemaXml: f.SchemaXml
        };

        if (f.TypeAsString === 'Lookup' || f.TypeAsString === 'LookupMulti') {
          const xml = f.SchemaXml;
          const listGuidMatch = xml.match(/List="\{?([0-9a-fA-F-]{36})\}?"/i);
          const showFieldMatch = xml.match(/ShowField="([^"]+)"/i);
          const webIdMatch = xml.match(/WebId="\{?([0-9a-fA-F-]{36})\}?"/i);

          if (listGuidMatch) {
            fieldObj.IsLookup = true;
            fieldObj.LookupField = showFieldMatch ? showFieldMatch[1] : 'Title';
            const lookupListId = listGuidMatch[1];

            if (webIdMatch && currentWebId && webIdMatch[1].toLowerCase() !== currentWebId.toLowerCase()) {
              fieldObj.CrossWeb = true;
            } else if (lookupListId.toLowerCase() === listId.toLowerCase()) {
              fieldObj.SelfReference = true;
            } else {
              try {
                const tlRes = await fetch(`${siteUrl}/_api/web/lists(guid'${lookupListId}')?$select=Title,Id`, { headers });
                if (tlRes.ok) {
                  const tl = (await tlRes.json()).d;
                  fieldObj.LookupListTitle = tl.Title;
                  if (tl.Title !== 'User Information List') {
                    await exportOneList(true, tl.Id);
                  }
                }
              } catch (e) {
                fieldObj.LookupResolutionFailed = true;
              }
            }
          }
        }

        fields.push(fieldObj);
      }

      let views = [];
      if (includeViews) {
        const viewsRes = await fetch(`${siteUrl}/_api/web/lists(guid'${listId}')/views`, { headers });
        if (viewsRes.ok) {
          const viewsData = (await viewsRes.json()).d;
          const rawViews = viewsData.results || viewsData;
          for (const v of rawViews) {
            if (v.PersonalView) continue;
            const vfRes = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listId}')/views/getbytitle('${enc(v.Title)}')/viewfields`,
              { headers }
            );
            let fieldNames = [];
            if (vfRes.ok) {
              const vfData = (await vfRes.json()).d;
              fieldNames = (vfData.Items && vfData.Items.results) || vfData.results || [];
            }
            views.push({
              Title: v.Title, DefaultView: v.DefaultView, Hidden: v.Hidden,
              Paged: v.Paged, RowLimit: v.RowLimit, ViewQuery: v.ViewQuery,
              ViewType: v.ViewType, ViewFields: fieldNames
            });
          }
        }
      }

      const schemaObj = {
        listName: listData.Title,
        description: listData.Description,
        templateType: listData.BaseTemplate,
        fields,
        views
      };

      cache.set(listId, schemaObj);
      inProgress.delete(listId);
      collected.push(schemaObj);
      return schemaObj;
    }

    const mainLists = [];
    for (const name of listNames) {
      const obj = await exportOneList(false, name);
      mainLists.push(obj.listName);
    }

    return {
      exportedFrom: siteUrl,
      exportedAt: new Date().toISOString(),
      mainLists,
      lists: collected
    };
  }

  // --------------------------------------------------------------------
  // IMPORT
  // --------------------------------------------------------------------
  async function importSchema(siteUrl, schema, includeViews) {
    siteUrl = siteUrl.replace(/\/$/, '');
    const logLines = [];
    const record = (msg, type = 'info') => {
      logLines.push({ msg, type });
      const prefix = type === 'warn' ? '⚠️' : type === 'ok' ? '✅' : 'ℹ️';
      console.log(`${prefix} ${msg}`);
    };

    const digestRes = await fetch(`${siteUrl}/_api/contextinfo`, { method: 'POST', headers });
    if (!digestRes.ok) throw new Error(`Could not get form digest (HTTP ${digestRes.status}). Check site URL / permissions.`);
    const digest = (await digestRes.json()).d.GetContextWebInformation.FormDigestValue;
    const writeHeaders = { ...headers, 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest };

    const lists = schema.lists || [schema];
    const mainListsArr = schema.mainLists || [lists[lists.length - 1].listName];
    const listMap = new Map();

    for (let i = 0; i < lists.length; i++) {
      const srcList = lists[i];
      const targetTitle = srcList.listName;
      const isDependencyOnly = !mainListsArr.includes(srcList.listName);

      let listId, createdListTitle, existed;
      const checkRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${enc(targetTitle)}')?$select=Id,Title`, { headers });
      existed = checkRes.ok;

      if (existed) {
        const d = (await checkRes.json()).d;
        listId = d.Id; createdListTitle = d.Title;
        record(`List "${createdListTitle}" already exists — syncing columns incrementally.`, 'info');
      } else {
        const createRes = await fetch(`${siteUrl}/_api/web/lists`, {
          method: 'POST',
          headers: writeHeaders,
          body: JSON.stringify({
            '__metadata': { type: 'SP.List' },
            Title: targetTitle,
            Description: srcList.description || '',
            BaseTemplate: srcList.templateType || 100
          })
        });
        if (!createRes.ok) {
          const err = await createRes.json().catch(() => null);
          record(`Failed to create list "${targetTitle}": ${err?.error?.message?.value || createRes.status}`, 'warn');
          listMap.set(srcList.listName, null);
          continue;
        }
        const d = (await createRes.json()).d;
        listId = d.Id; createdListTitle = d.Title;
        record(`List "${createdListTitle}" created${isDependencyOnly ? ' (lookup dependency)' : ''}.`, 'ok');
      }
      listMap.set(srcList.listName, { id: listId, title: createdListTitle });

      const existingFieldsRes = await fetch(
        `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,TypeAsString,Required`,
        { headers }
      );
      const existingFields = existingFieldsRes.ok ? ((await existingFieldsRes.json()).d.results || []) : [];
      const existingMap = new Map(existingFields.map(f => [f.InternalName, f]));

      // A newly created SharePoint list already has its built-in Title field.
      // Fields are matched by InternalName, so the native Title field is
      // synced (not duplicated) rather than re-created.
      async function syncRequiredSetting(field) {
        if (typeof field.Required !== 'boolean') return true;

        const safeInternalName = enc(String(field.InternalName).replace(/'/g, "''"));
        const updateRes = await fetch(
          `${siteUrl}/_api/web/lists(guid'${listId}')/fields/getbyinternalnameortitle('${safeInternalName}')`,
          {
            method: 'POST',
            headers: { ...writeHeaders, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
            body: JSON.stringify({
              '__metadata': { type: 'SP.Field' },
              Required: field.Required
            })
          }
        );

        if (!updateRes.ok) {
          const err = await updateRes.json().catch(() => null);
          record(`[${createdListTitle}] Could not update required setting for "${field.Title}": ${err?.error?.message?.value || updateRes.status}`, 'warn');
          return false;
        }
        return true;
      }

      for (const field of (srcList.fields || [])) {
        let schemaXml = field.SchemaXml;

        if (field.IsLookup) {
          if (field.CrossWeb) {
            record(`[${createdListTitle}] Skipped "${field.Title}": looks up to a list in a different site (sub-web) — not auto-resolved.`, 'warn');
            continue;
          }
          let destListId = null;
          if (field.SelfReference) {
            destListId = listId;
          } else if (field.LookupListTitle) {
            const mapped = listMap.get(field.LookupListTitle);
            if (mapped) {
              destListId = mapped.id;
            } else {
              const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${enc(field.LookupListTitle)}')?$select=Id`, { headers });
              if (r.ok) destListId = (await r.json()).d.Id;
            }
          }
          if (!destListId) {
            record(`[${createdListTitle}] Skipped "${field.Title}": couldn't resolve lookup target "${field.LookupListTitle}" (possibly circular) — add manually.`, 'warn');
            continue;
          }
          schemaXml = schemaXml.replace(/List="\{?[0-9a-fA-F-]{36}\}?"/i, `List="{${destListId}}"`);
        }

        const existing = existingMap.get(field.InternalName);

        if (!existing) {
          try {
            const res = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listId}')/fields/createfieldasxml`,
              {
                method: 'POST',
                headers: writeHeaders,
                body: JSON.stringify({
                  parameters: {
                    '__metadata': { type: 'SP.XmlSchemaFieldCreationInformation' },
                    SchemaXml: schemaXml,
                    Options: 8
                  }
                })
              }
            );
            if (res.ok) {
              await syncRequiredSetting(field);
              record(`[${createdListTitle}] Field added: ${field.Title} (${field.InternalName})`, 'ok');
            } else {
              const err = await res.json().catch(() => null);
              record(`[${createdListTitle}] Field failed: ${field.Title} — ${err?.error?.message?.value || res.status}`, 'warn');
            }
          } catch (e) {
            record(`[${createdListTitle}] Field error: ${field.Title} — ${e.message}`, 'warn');
          }
        } else if (existing.TypeAsString !== field.TypeAsString) {
          record(`[${createdListTitle}] Type mismatch on "${field.InternalName}": source=${field.TypeAsString}, destination=${existing.TypeAsString}. Not auto-changed — SharePoint can't safely change a column's type via API. Review manually.`, 'warn');
        } else {
          if (typeof field.Required === 'boolean' && existing.Required !== field.Required) {
            if (await syncRequiredSetting(field)) {
              record(`[${createdListTitle}] Required setting updated: ${field.InternalName} → ${field.Required ? 'required' : 'optional'}`, 'ok');
            }
          } else {
            record(`[${createdListTitle}] Field already present, type and required setting match: ${field.InternalName}`, 'info');
          }
        }
      }

      if (includeViews && srcList.views && srcList.views.length) {
        for (const view of srcList.views) {
          const checkViewRes = await fetch(
            `${siteUrl}/_api/web/lists(guid'${listId}')/views/getbytitle('${enc(view.Title)}')`,
            { headers }
          );
          const viewExists = checkViewRes.ok;

          try {
            if (!viewExists) {
              const createViewRes = await fetch(
                `${siteUrl}/_api/web/lists(guid'${listId}')/views`,
                {
                  method: 'POST',
                  headers: writeHeaders,
                  body: JSON.stringify({
                    '__metadata': { type: 'SP.View' },
                    Title: view.Title,
                    PersonalView: false,
                    RowLimit: view.RowLimit || 30,
                    Paged: view.Paged !== false
                  })
                }
              );
              if (!createViewRes.ok) {
                const err = await createViewRes.json().catch(() => null);
                record(`[${createdListTitle}] View failed: ${view.Title} — ${err?.error?.message?.value || createViewRes.status}`, 'warn');
                continue;
              }
            }

            // Replace the complete View XML (not addviewfield) so that a new
            // view's default Title/LinkTitle entry doesn't cause Title to
            // appear twice in the target view.
            const uniqueViewFields = [...new Set((view.ViewFields || []).filter(Boolean))];
            const viewFieldsXml = uniqueViewFields.map(name => `<FieldRef Name="${name}" />`).join('');
            const queryXml = view.ViewQuery ? `<Query>${view.ViewQuery}</Query>` : '';
            const viewXml = `<View><ViewFields>${viewFieldsXml}</ViewFields>${queryXml}<RowLimit Paged="${view.Paged !== false}">${view.RowLimit || 30}</RowLimit></View>`;

            const setViewRes = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listId}')/views/getbytitle('${enc(view.Title)}')/SetViewXml`,
              {
                method: 'POST',
                headers: writeHeaders,
                body: JSON.stringify({ viewXml })
              }
            );
            if (!setViewRes.ok) {
              const err = await setViewRes.json().catch(() => null);
              record(`[${createdListTitle}] View fields failed: ${view.Title} — ${err?.error?.message?.value || setViewRes.status}`, 'warn');
              continue;
            }

            record(`[${createdListTitle}] View ${viewExists ? 'updated' : 'created'}: ${view.Title}`, 'ok');
          } catch (e) {
            record(`[${createdListTitle}] View error: ${view.Title} — ${e.message}`, 'warn');
          }
        }
      }
    }

    const doneUrls = mainListsArr
      .map(name => listMap.get(name))
      .filter(Boolean)
      .map(m => `${siteUrl}/Lists/${enc(m.title)}`);

    if (doneUrls.length) {
      record(`Done. ${doneUrls.length} main list(s):`, 'info');
      doneUrls.forEach(u => record(`  ${u}`, 'info'));
    } else {
      record(`Done — see warnings above.`, 'warn');
    }
    return logLines;
  }

  // --------------------------------------------------------------------
  // PUBLIC API — exposed on window for console use
  // --------------------------------------------------------------------
  window.SPMigrator = {
    /** List all non-system lists on a site. */
    async listLists(siteUrl = currentSiteUrl()) {
      const lists = await getLists(siteUrl);
      console.table(lists);
      return lists;
    },

    /**
     * Export one or more lists (with lookup dependencies and optionally views)
     * to a downloaded JSON file.
     * Example: SPMigrator.export(['Tasks', 'Projects'], true)
     */
    async export(listNames, includeViews = true, siteUrl = currentSiteUrl()) {
      if (!Array.isArray(listNames) || listNames.length === 0) {
        throw new Error('Provide an array of list names, e.g. SPMigrator.export(["Tasks"])');
      }
      console.log(`Exporting ${listNames.length} list(s) from ${siteUrl} ...`);
      const schema = await exportMultipleSchemas(siteUrl, listNames, includeViews);
      const filenameBase = listNames.length === 1 ? listNames[0] : `${listNames.length}-lists`;
      const filename = `${filenameBase.replace(/[^a-z0-9\-_]+/gi, '_')}-schema.json`;
      downloadJson(schema, filename);
      const depCount = schema.lists.length - schema.mainLists.length;
      console.log(`✅ Exported ${schema.mainLists.length} selected list(s) plus ${depCount} lookup-dependency list(s) (${schema.lists.length} total). File downloaded: ${filename}`);
      return schema;
    },

    /**
     * Import a schema (previously exported) into the given destination site.
     * Opens a file picker to select the schema JSON.
     * Example: await SPMigrator.import(true)
     */
    async import(includeViews = true, siteUrl = currentSiteUrl()) {
      console.log('Select the schema JSON file to import...');
      const schema = await pickJsonFile();
      console.log(`Processing ${siteUrl} ...`);
      const logLines = await importSchema(siteUrl, schema, includeViews);
      return logLines;
    },

    /**
     * Import a schema object you already have in a variable (no file picker).
     * Example: await SPMigrator.importSchemaObject(mySchemaVar, true)
     */
    async importSchemaObject(schema, includeViews = true, siteUrl = currentSiteUrl()) {
      console.log(`Processing ${siteUrl} ...`);
      const logLines = await importSchema(siteUrl, schema, includeViews);
      return logLines;
    }
  };

  console.log('%cSPMigrator loaded.', 'font-weight:bold;color:#0a7;');
  console.log('Usage:');
  console.log('  await SPMigrator.listLists()                      // list all lists on this site');
  console.log('  await SPMigrator.export(["Tasks","Projects"])     // export lists, downloads JSON');
  console.log('  await SPMigrator.import()                         // pick JSON file, import to this site');
})();