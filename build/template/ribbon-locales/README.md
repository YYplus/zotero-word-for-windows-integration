# Word Ribbon localizations

Ribbon translations live in one JSON file per locale. The English strings in `customUI.xml` are the built-in fallback, so no `en-US.json` file is required.

Each locale file contains:

- `locale`: the locale identifier used in the localization store
- `officeLanguageIDs`: one or more Office UI language IDs that should use this locale
- `controls`: translated `label` and `supertip` strings keyed by Ribbon control ID

The generator uses only the Python 3 standard library. It validates locale identifiers, duplicate JSON keys, Office language ID conflicts, control coverage, English fallback tags, and XML 1.0-safe translation text.

## First-time integration or VBA changes

When the localization callbacks in `ZoteroRibbon.bas` are first introduced, or whenever VBA code changes:

1. Open `install/Zotero.dotm` in the oldest Word version supported by the project.
2. Import or replace the updated VBA module, then use **Debug -> Compile Project** and save the template.
3. Run:

   ```sh
   python3 build/template/generate_ribbon_localizations.py
   ```

   This creates or updates the Ribbon localization Custom XML data, its relationships and content type, and synchronizes `customUI.xml` inside `install/Zotero.dotm`.
4. Run the normal template source synchronization/checks, including `build/template/unpack_templates.sh`.
5. Test the resulting template through Word's normal STARTUP add-in loading path.

The generator does not compile or validate VBA bytecode. A successful generator run or `--check` therefore does not replace the Word compile step when VBA has changed.

## Adding or updating a locale

Once the base template already contains the generic localization callbacks:

1. Add or edit `<locale>.json` in this directory.
2. Run:

   ```sh
   python3 build/template/generate_ribbon_localizations.py
   ```

3. Run the normal template source synchronization/checks.

No VBA or Ribbon XML changes are required for locale-only updates.

At runtime, Word's UI language ID is resolved through the generated Custom XML store. If no locale is mapped, the localization store is unavailable, a query fails, or a translated string cannot be found, the callbacks keep the English strings stored in each Ribbon control's `tag` attribute.

Use:

```sh
python3 build/template/generate_ribbon_localizations.py --check
```

to verify the generated localization sources and package metadata. XML comparisons are structural, so formatting changes made by `xmllint` during normal template unpacking do not cause false failures. The check also verifies the expected Custom XML properties/relationships/content types and rejects duplicate ZIP entries, but it intentionally does not claim to validate compiled VBA bytecode.
