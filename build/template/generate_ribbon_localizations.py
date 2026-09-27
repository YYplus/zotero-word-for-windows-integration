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
DATA_STORE_ITEM_ID = "{6F533F15-7F4D-47D7-96DF-919E748E3D85}"

CUSTOM_UI_NS = "http://schemas.microsoft.com/office/2006/01/customui"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DATA_STORE_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"

CUSTOM_XML_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)
CUSTOM_XML_PROPS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)
CUSTOM_XML_PROPS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
)

CUSTOM_XML_PART = "customXml/item1.xml"
CUSTOM_XML_PROPS_PART = "customXml/itemProps1.xml"
CUSTOM_XML_RELS_PART = "customXml/_rels/item1.xml.rels"
DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"
CONTENT_TYPES_PART = "[Content_Types].xml"
CUSTOM_UI_PART = "customUI/customUI.xml"

REPO_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = Path(__file__).resolve().parent
LOCALES_DIR = TEMPLATE_DIR / "ribbon-locales"
SOURCE_TEMPLATE = TEMPLATE_DIR / "Zotero.dotm"
SOURCE_CUSTOM_UI = SOURCE_TEMPLATE / CUSTOM_UI_PART
SOURCE_CUSTOM_XML = SOURCE_TEMPLATE / CUSTOM_XML_PART
SOURCE_CUSTOM_XML_PROPS = SOURCE_TEMPLATE / CUSTOM_XML_PROPS_PART
SOURCE_CUSTOM_XML_RELS = SOURCE_TEMPLATE / CUSTOM_XML_RELS_PART
INSTALL_TEMPLATE = REPO_ROOT / "install" / "Zotero.dotm"


def serialize(root):
    if hasattr(ET, "indent"):
        ET.indent(root, space="  ")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True) + b"\n"


def get_ribbon_controls():
    root = ET.parse(SOURCE_CUSTOM_UI).getroot()
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
    files = sorted(LOCALES_DIR.glob("*.json"))
    if not files:
        raise ValueError(f"No locale files found in {LOCALES_DIR}")

    expected = set(control_ids)
    locales = []
    seen_locales = set()
    seen_language_ids = {}

    for path in files:
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

        if not isinstance(controls, dict) or set(controls) != expected:
            raise ValueError(
                f"{path.name}: controls must match the localized controls in customUI.xml"
            )
        for control_id in control_ids:
            strings = controls[control_id]
            if not isinstance(strings, dict):
                raise ValueError(f"{path.name}: {control_id} must be an object")
            for key in ("label", "supertip"):
                if not isinstance(strings.get(key), str) or not strings[key]:
                    raise ValueError(
                        f"{path.name}: {control_id}.{key} must be non-empty"
                    )

        locales.append(data)

    return locales


def build_localization_xml(locales, control_ids):
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
            control = ET.SubElement(
                locale_node,
                f"{{{RIBBON_NAMESPACE}}}control",
                {"id": control_id},
            )
            label = ET.SubElement(control, f"{{{RIBBON_NAMESPACE}}}label")
            label.text = data["controls"][control_id]["label"]
            supertip = ET.SubElement(control, f"{{{RIBBON_NAMESPACE}}}supertip")
            supertip.text = data["controls"][control_id]["supertip"]

    return serialize(root)


def build_item_props():
    ET.register_namespace("ds", DATA_STORE_NS)
    root = ET.Element(
        f"{{{DATA_STORE_NS}}}dataStoreItem",
        {f"{{{DATA_STORE_NS}}}itemID": DATA_STORE_ITEM_ID},
    )
    refs = ET.SubElement(root, f"{{{DATA_STORE_NS}}}schemaRefs")
    ET.SubElement(
        refs,
        f"{{{DATA_STORE_NS}}}schemaRef",
        {f"{{{DATA_STORE_NS}}}uri": RIBBON_NAMESPACE},
    )
    return serialize(root)


def build_item_rels():
    ET.register_namespace("", REL_NS)
    root = ET.Element(f"{{{REL_NS}}}Relationships")
    ET.SubElement(
        root,
        f"{{{REL_NS}}}Relationship",
        {
            "Id": "rId1",
            "Type": CUSTOM_XML_PROPS_REL_TYPE,
            "Target": "itemProps1.xml",
        },
    )
    return serialize(root)


def update_document_rels(data):
    root = ET.fromstring(data)
    target = "../" + CUSTOM_XML_PART
    relationships = root.findall(f"{{{REL_NS}}}Relationship")

    for rel in relationships:
        if rel.get("Type") == CUSTOM_XML_REL_TYPE and rel.get("Target") == target:
            return serialize(root)

    relationship_id = "rIdZoteroRibbonLocalization"
    existing_ids = {rel.get("Id") for rel in relationships}
    suffix = 1
    while relationship_id in existing_ids:
        relationship_id = f"rIdZoteroRibbonLocalization{suffix}"
        suffix += 1

    ET.SubElement(
        root,
        f"{{{REL_NS}}}Relationship",
        {
            "Id": relationship_id,
            "Type": CUSTOM_XML_REL_TYPE,
            "Target": target,
        },
    )
    return serialize(root)


def update_content_types(data):
    root = ET.fromstring(data)
    part_name = "/" + CUSTOM_XML_PROPS_PART

    has_xml_default = any(
        node.get("Extension", "").lower() == "xml"
        for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
    )
    if not has_xml_default:
        ET.SubElement(
            root,
            f"{{{CONTENT_TYPES_NS}}}Default",
            {"Extension": "xml", "ContentType": "application/xml"},
        )

    for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if node.get("PartName") == part_name:
            node.set("ContentType", CUSTOM_XML_PROPS_CONTENT_TYPE)
            return serialize(root)

    ET.SubElement(
        root,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {
            "PartName": part_name,
            "ContentType": CUSTOM_XML_PROPS_CONTENT_TYPE,
        },
    )
    return serialize(root)


def patch_template(custom_ui, localization_xml, item_props, item_rels):
    if not INSTALL_TEMPLATE.exists():
        raise FileNotFoundError(f"Template not found: {INSTALL_TEMPLATE}")

    with zipfile.ZipFile(INSTALL_TEMPLATE, "r") as src:
        names = set(src.namelist())
        for required in (DOCUMENT_RELS_PART, CONTENT_TYPES_PART):
            if required not in names:
                raise RuntimeError(
                    f"{INSTALL_TEMPLATE}: required package part {required} is missing"
                )

        replacements = {
            CUSTOM_UI_PART: custom_ui,
            CUSTOM_XML_PART: localization_xml,
            CUSTOM_XML_PROPS_PART: item_props,
            CUSTOM_XML_RELS_PART: item_rels,
            DOCUMENT_RELS_PART: update_document_rels(src.read(DOCUMENT_RELS_PART)),
            CONTENT_TYPES_PART: update_content_types(src.read(CONTENT_TYPES_PART)),
        }

        fd, temp_name = tempfile.mkstemp(
            prefix="Zotero-ribbon-locales-",
            suffix=".dotm",
            dir=str(INSTALL_TEMPLATE.parent),
        )
        os.close(fd)
        temp_path = Path(temp_name)

        try:
            with zipfile.ZipFile(temp_path, "w") as dst:
                dst.comment = src.comment
                written = set()
                for info in src.infolist():
                    if info.filename in replacements:
                        dst.writestr(info, replacements[info.filename])
                        written.add(info.filename)
                    else:
                        dst.writestr(info, src.read(info.filename))
                for name, data in replacements.items():
                    if name not in written:
                        dst.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
            shutil.move(str(temp_path), str(INSTALL_TEMPLATE))
        finally:
            if temp_path.exists():
                temp_path.unlink()


def write_sources(localization_xml, item_props, item_rels):
    generated = {
        SOURCE_CUSTOM_XML: localization_xml,
        SOURCE_CUSTOM_XML_PROPS: item_props,
        SOURCE_CUSTOM_XML_RELS: item_rels,
    }
    for path, data in generated.items():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)


def check(localization_xml, item_props, item_rels):
    expected_sources = {
        SOURCE_CUSTOM_XML: localization_xml,
        SOURCE_CUSTOM_XML_PROPS: item_props,
        SOURCE_CUSTOM_XML_RELS: item_rels,
    }
    errors = []
    for path, expected in expected_sources.items():
        if not path.exists() or path.read_bytes() != expected:
            errors.append(f"out of date or missing: {path.relative_to(REPO_ROOT)}")

    if not INSTALL_TEMPLATE.exists():
        errors.append(f"missing: {INSTALL_TEMPLATE.relative_to(REPO_ROOT)}")
        return errors

    with zipfile.ZipFile(INSTALL_TEMPLATE, "r") as zf:
        names = set(zf.namelist())
        required = (
            CUSTOM_UI_PART,
            CUSTOM_XML_PART,
            CUSTOM_XML_PROPS_PART,
            CUSTOM_XML_RELS_PART,
            DOCUMENT_RELS_PART,
            CONTENT_TYPES_PART,
        )
        for name in required:
            if name not in names:
                errors.append(f"install/Zotero.dotm is missing {name}")

        if CUSTOM_UI_PART in names and zf.read(CUSTOM_UI_PART) != SOURCE_CUSTOM_UI.read_bytes():
            errors.append("install/Zotero.dotm has out-of-date customUI.xml")
        if CUSTOM_XML_PART in names and zf.read(CUSTOM_XML_PART) != localization_xml:
            errors.append("install/Zotero.dotm has out-of-date Ribbon localization XML")

        if DOCUMENT_RELS_PART in names:
            rels = ET.fromstring(zf.read(DOCUMENT_RELS_PART))
            target = "../" + CUSTOM_XML_PART
            if not any(
                rel.get("Type") == CUSTOM_XML_REL_TYPE and rel.get("Target") == target
                for rel in rels.findall(f"{{{REL_NS}}}Relationship")
            ):
                errors.append("install/Zotero.dotm is missing the Ribbon localization relationship")

        if CONTENT_TYPES_PART in names:
            types = ET.fromstring(zf.read(CONTENT_TYPES_PART))
            part_name = "/" + CUSTOM_XML_PROPS_PART
            if not any(
                node.get("PartName") == part_name
                and node.get("ContentType") == CUSTOM_XML_PROPS_CONTENT_TYPE
                for node in types.findall(f"{{{CONTENT_TYPES_NS}}}Override")
            ):
                errors.append(
                    "install/Zotero.dotm is missing the custom XML properties content type"
                )

    return errors


def main():
    parser = argparse.ArgumentParser(
        description="Build locale files into the Zotero Word Ribbon localization store."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Check generated sources and Zotero.dotm without modifying files.",
    )
    args = parser.parse_args()

    control_ids = get_ribbon_controls()
    locales = load_locales(control_ids)
    localization_xml = build_localization_xml(locales, control_ids)
    item_props = build_item_props()
    item_rels = build_item_rels()
    custom_ui = SOURCE_CUSTOM_UI.read_bytes()

    if args.check:
        errors = check(localization_xml, item_props, item_rels)
        if errors:
            for error in errors:
                print(error)
            raise SystemExit(1)
        print("Ribbon localizations are up to date.")
        return

    write_sources(localization_xml, item_props, item_rels)
    patch_template(custom_ui, localization_xml, item_props, item_rels)
    print(f"Built {len(locales)} Ribbon locale(s) into install/Zotero.dotm.")


if __name__ == "__main__":
    main()
