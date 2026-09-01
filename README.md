# SharePoint List Migrator (Browser Console Script)

A standalone JavaScript script to export list schemas (columns, views, lookup
dependencies, required/optional settings) from one SharePoint site and import
them into another — no installation required, runs directly in the browser
console.

## How to Use

### 1. Load the script
1. Open the source SharePoint site in your browser.
2. Open DevTools Console (F12 → Console tab, or Ctrl+Shift+J).
3. Copy the entire contents of `sp-list-migrator.js` and paste it into the console, then press Enter.
4. You should see: `SPMigrator loaded.`

### 2. List available lists (source site)
```js
await SPMigrator.listLists();
```
Prints a table of all non-system lists with item counts and types.

### 3. Get list names as a copy-paste-ready array
```js
console.log(JSON.stringify((await SPMigrator.listLists()).map(l => l.Title), null, 2));
```
Copy the output, delete any list names you don't want to migrate, and use the remaining array in the next step.

### 4. Export list(s)
```js
await SPMigrator.export(["Tasks", "Projects"]);
```
- Second argument (`includeViews`, default `true`) controls whether views are exported.
- Automatically resolves and includes lookup-column dependencies (lists the lookups point to).
- Downloads a `*-schema.json` file to your Downloads folder.

### 5. Import into destination site
1. Navigate to the **destination** SharePoint site.
2. Open DevTools Console and paste the script again.
3. Run:
```js
await SPMigrator.import();
```
4. A file picker will open — select the JSON file downloaded in step 4.
5. The script creates missing lists/columns/views, syncs Required/Optional settings on existing fields, and resolves lookup column targets automatically.

### Alternative: import from a variable (no file picker)
```js
await SPMigrator.importSchemaObject(mySchemaObject, true);
```

## Notes
- Uses `application/json;odata=verbose` REST calls against `_api/web/lists`.
- Skips hidden/system fields and noisy default system lists.
- Cross-site (sub-web) lookup columns are flagged as warnings and not auto-resolved — add manually.
- Column type changes on existing fields are flagged but not auto-applied (SharePoint REST API cannot safely change a field's type).

## Requirements
- Modern browser (Chrome/Edge recommended).
- Site Owner/Member permissions on both source and destination sites.
```
