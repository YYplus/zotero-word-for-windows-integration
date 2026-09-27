# Word Ribbon localizations

Ribbon translations live in one JSON file per locale. The English strings in `customUI.xml` are the built-in fallback, so no `en-US.json` file is required.

Each locale file contains:

- `locale`: the locale identifier used in the localization store
- `officeLanguageIDs`: one or more Office UI language IDs that should use this locale
- `controls`: translated `label` and `supertip` strings keyed by Ribbon control ID

To add or update a locale:

1. Add or edit `<locale>.json` in this directory.
2. Run:

   ```sh
   python3 build/template/generate_ribbon_localizations.py
   ```

   The generator rebuilds the Ribbon localization custom XML files and creates or updates the corresponding custom XML parts, relationships, content type, and `customUI.xml` inside `install/Zotero.dotm`.
3. Run the normal template source synchronization/checks.

No VBA changes are required when a locale is added. The generator validates that every locale covers the same localized Ribbon controls and that an Office language ID is not assigned to more than one locale.

At runtime, Word's UI language ID is resolved through the generated custom XML store. If no locale is mapped, the localization store is unavailable, or a translated string cannot be found, the callbacks fall back to the English strings stored in each Ribbon control's `tag` attribute.

Use `python3 build/template/generate_ribbon_localizations.py --check` to verify that the generated source files and the localization data embedded in `install/Zotero.dotm` are in sync.
