#!/usr/bin/env python3
import argparse
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

RIBBON_NAMESPACE = "http://www.zotero.org/ribbon-localization/v1"
DEFAULT_LOCALE = "en-US"

REPO_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = Path(__file__).resolve().parent
LOCALES_DIR = TEMPLATE_DIR / "ribbon-locales"
CUSTOM_UI = TEMPLATE_DIR / "Zotero.dotm" / "customUI" / "customUI.xml"
SOURCE_LOCALIZATIONS = TEMPLATE_DIR / "Zotero.dotm" / "customXml" / "item1.xml"
INSTALL_TEMPLATE = REPO_ROOT / "install" / "Zotero.dotm"
PACKAGE_LOCALIZATIONS = "customXml/item1.xml"

CUSTOM_UI_NS = "http://schemas.microsoft.com/office/2006/01/customui"


def get_ribbon_controls():
    root = ET.parse(CUSTOM_UI).getroot()
    controls = []
    for button in root.findall(f".//{{{CUSTOM_UI_NS}}}button"):
        if (
            button.get("getLabel") == "ZoteroRibbon.ZoteroRibbonGetLabel"
            and button.get("getSupertip") == "ZoteroRibbon.ZoteroRibbonGetSupertip"
        ):
            control_id = button.get("id")
            if not control_id:
                raise ValueError("Localized Ribbon buttons must have an id")
            controls.append(control_id)
    if not controls:
        raise ValueError("No localized Ribbon controls found in customUI.xml")
    return controls


def load_locales(control_ids):
    locale_files = sorted(LOCALES_DIR.glob("*.json"))
    if not locale_files:
        raise ValueError(f"No locale files found in {LOCALES_DIR}")

    expected_controls = set(control_ids)
    locales = []
    seen_locales = set()
    seen_language_ids = {}

    for path in locale_files:
        data = json.loads(path.read_text(encoding="utf-8"))
        locale = data.get("locale")
        language_ids = data.get("officeLanguageIDs")
        controls = data.get("controls")

        if not isinstance(locale, str) or not locale:
            raise ValueError(f"{path.name}: locale must be a non-empty string")
        if path.stem != locale:
            raise ValueError(f"{path.name}: filename must be {locale}.json")
        if locale == DEFAULT_LOCALE:
            raise ValueError(
                f"{DEFAULT_LOCALE} is the built-in fallback and must not have a locale file"
            )
        if locale in seen_locales:
            raise ValueError(f"Duplicate locale: {locale}")
        seen_locales.add(locale)

        if not isinstance(language_ids, list) or not language_ids:
            raise ValueError(
                f"{path.name}: officeLanguageIDs must be a non-empty list"
            )
        for language_id in language_ids:
            if not isinstance(language_id, int) or language_id <= 0:
                raise ValueError(
                    f"{path.name}: invalid Office language ID {language_id!r}"
                )
            if language_id in seen_language_ids:
                raise ValueError(
                    f"Office language ID {language_id} is mapped by both "
                    f"{seen_language_ids[language_id]} and {locale}"
                )
            seen_language_ids[language_id] = locale

        if not isinstance(controls, dict):
            raise ValueError(f"{path.name}: controls must be an object")
        actual_controls = set(controls)
        if actual_controls != expected_controls:
            missing = sorted(expected_controls - actual_controls)
            extra = sorted(actual_controls - expected_controls)
            details = []
            if missing:
                details.append("missing " + ", ".join(missing))
            if extra:
                details.append("extra " + ", ".join(extra))
            raise ValueError(f"{path.name}: control mismatch ({'; '.join(details)})")

        for control_id in control_ids:
            strings = controls[control_id]
            if not isinstance(strings, dict):
                raise ValueError(f"{path.name}: {control_id} must be an object")
            for key in ("label", "supertip"):
                value = strings.get(key)
                if not isinstance(value, str) or not value:
                    raise ValueError(
                        f"{path.name}: {control_id}.{key} must be non-empty"
                    )

        locales.append(data)

    return locales


def build_xml(locales, control_ids):
    ET.register_namespace("", RIBBON_NAMESPACE)
    root = ET.Element(
        f"{{{RIBBON_NAMESPACE}}}ribbonLocalization",
        {"defaultLocale": DEFAULT_LOCALE},
    )

    languages = ET.SubElement(root, f"{{{RIBBON_NAMESPACE}}}languages")
    for data in locales:
        for language_id in data["officeLanguageIDs"]:
            ET.SubElement(
                languages,
                f"{{{RIBBON_NAMESPACE}}}language",
                {
                    "officeLanguageID": str(language_id),
                    "locale": data["locale"],
                },
            )

    locale_root = ET.SubElement(root, f"{{{RIBBON_NAMESPACE}}}locales")
    for data in locales:
        locale_node = ET.SubElement(
            locale_root,
            f"{{{RIBBON_NAMESPACE}}}locale",
            {"id": data["locale"]},
        )
        for control_id in control_ids:
            control_node = ET.SubElement(
                locale_node,
                f"{{{RIBBON_NAMESPACE}}}control",
                {"id": control_id},
            )
            label = ET.SubElement(control_node, f"{{{RIBBON_NAMESPACE}}}label")
            label.text = data["controls"][control_id]["label"]
            supertip = ET.SubElement(
                control_node, f"{{{RIBBON_NAMESPACE}}}supertip"
            )
            supertip.text = data["controls"][control_id]["supertip"]

    if hasattr(ET, "indent"):
        ET.indent(root, space="  ")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True) + b"\n"


def replace_package_part(package_path, part_name, data):
    fd, temp_name = tempfile.mkstemp(
        prefix="Zotero-ribbon-locales-",
        suffix=".dotm",
        dir=str(package_path.parent),
    )
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        found = False
        with zipfile.ZipFile(package_path, "r") as src, zipfile.ZipFile(
            temp_path, "w"
        ) as dst:
            dst.comment = src.comment
            for info in src.infolist():
                if info.filename == part_name:
                    dst.writestr(info, data)
                    found = True
                else:
                    dst.writestr(info, src.read(info.filename))
        if not found:
            raise RuntimeError(
                f"{package_path}: missing {part_name}; bootstrap the localization part first"
            )
        shutil.move(str(temp_path), str(package_path))
    finally:
        if temp_path.exists():
            temp_path.unlink()


def main():
    parser = argparse.ArgumentParser(
        description="Build locale files into the Zotero Word Ribbon localization part."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Check generated source and Zotero.dotm without modifying files.",
    )
    args = parser.parse_args()

    control_ids = get_ribbon_controls()
    locales = load_locales(control_ids)
    generated = build_xml(locales, control_ids)

    if args.check:
        errors = []
        if not SOURCE_LOCALIZATIONS.exists():
            errors.append(f"missing {SOURCE_LOCALIZATIONS.relative_to(REPO_ROOT)}")
        elif SOURCE_LOCALIZATIONS.read_bytes() != generated:
            errors.append(
                f"out of date: {SOURCE_LOCALIZATIONS.relative_to(REPO_ROOT)}"
            )

        if not INSTALL_TEMPLATE.exists():
            errors.append(f"missing {INSTALL_TEMPLATE.relative_to(REPO_ROOT)}")
        else:
            with zipfile.ZipFile(INSTALL_TEMPLATE, "r") as zf:
                try:
                    packaged = zf.read(PACKAGE_LOCALIZATIONS)
                except KeyError:
                    packaged = None
                if packaged != generated:
                    errors.append(
                        f"out of date or missing: "
                        f"{INSTALL_TEMPLATE.relative_to(REPO_ROOT)}/{PACKAGE_LOCALIZATIONS}"
                    )

        if errors:
            for error in errors:
                print(error)
            raise SystemExit(1)

        print("Ribbon localizations are up to date.")
        return

    SOURCE_LOCALIZATIONS.parent.mkdir(parents=True, exist_ok=True)
    SOURCE_LOCALIZATIONS.write_bytes(generated)
    replace_package_part(INSTALL_TEMPLATE, PACKAGE_LOCALIZATIONS, generated)
    print(f"Built {len(locales)} Ribbon locale(s) into install/Zotero.dotm.")


if __name__ == "__main__":
    main()
