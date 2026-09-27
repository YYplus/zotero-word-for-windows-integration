#!/usr/bin/env python3
import argparse
from collections import Counter
import json
import os
import re
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
VBA_PROJECT_PART = "word/vbaProject.bin"

REPO_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = Path(__file__).resolve().parent
LOCALES_DIR = TEMPLATE_DIR / "ribbon-locales"
SOURCE_TEMPLATE = TEMPLATE_DIR / "Zotero.dotm"
SOURCE_CUSTOM_UI = SOURCE_TEMPLATE / CUSTOM_UI_PART
SOURCE_CUSTOM_XML = SOURCE_TEMPLATE / CUSTOM_XML_PART
SOURCE_CUSTOM_XML_PROPS = SOURCE_TEMPLATE / CUSTOM_XML_PROPS_PART
SOURCE_CUSTOM_XML_RELS = SOURCE_TEMPLATE / CUSTOM_XML_RELS_PART
SOURCE_VBA = SOURCE_TEMPLATE / "word/vbaProject.bin/ZoteroRibbon.bas"
INSTALL_TEMPLATE = REPO_ROOT / "install" / "Zotero.dotm"

LOCALE_RE = re.compile(r"^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*$")
REQUIRED_VBA_CALLBACKS = (
    "Sub ZoteroRibbonGetLabel(",
    "Sub ZoteroRibbonGetSupertip(",
)


def indent_xml(element, level=0):
    indentation = "\n" + level * "  "
    child_indentation = "\n" + (level + 1) * "  "
    children = list(element)
    if children:
        if not element.text or not element.text.strip():
            element.text = child_indentation
        for child in children:
            indent_xml(child, level + 1)
        if not children[-1].tail or not children[-1].tail.strip():
            children[-1].tail = indentation
    if level and (not element.tail or not element.tail.strip()):
        element.tail = indentation


def serialize(root):
    indent_xml(root)
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>\n'
        + ET.tostring(root, encoding="utf-8")
        + b"\n"
    )


def parse_xml(data, description):
    try:
        return ET.fromstring(data)
    except (ET.ParseError, ValueError) as exc:
        raise ValueError(f"{description}: invalid XML: {exc}") from exc


def normalized_xml_node(node):
    text = node.text
    if text is not None and not text.strip():
        text = None
    tail = node.tail
    if tail is not None and not tail.strip():
        tail = None
    return (
        node.tag,
        tuple(sorted(node.attrib.items())),
        text,
        tail,
        tuple(normalized_xml_node(child) for child in node),
    )


def xml_semantically_equal(actual, expected):
    try:
        return normalized_xml_node(ET.fromstring(actual)) == normalized_xml_node(
            ET.fromstring(expected)
        )
    except (ET.ParseError, ValueError):
        return False


def validate_xml_text(value, context):
    for character in value:
        codepoint = ord(character)
        if not (
            codepoint in (0x9, 0xA, 0xD)
            or 0x20 <= codepoint <= 0xD7FF
            or 0xE000 <= codepoint <= 0xFFFD
            or 0x10000 <= codepoint <= 0x10FFFF
        ):
            raise ValueError(
                f"{context}: contains XML 1.0-invalid character U+{codepoint:04X}"
            )


def load_json(path):
    def no_duplicate_keys(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"{path.name}: duplicate JSON key {key!r}")
            result[key] = value
        return result

    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=no_duplicate_keys,
        )
    except json.JSONDecodeError as exc:
        raise ValueError(f"{path.name}: invalid JSON: {exc}") from exc


def get_ribbon_controls():
    root = ET.parse(SOURCE_CUSTOM_UI).getroot()
    controls = []
    seen_ids = set()

    for button in root.findall(f".//{{{CUSTOM_UI_NS}}}button"):
        if (
            button.get("getLabel") == "ZoteroRibbon.ZoteroRibbonGetLabel"
            and button.get("getSupertip") == "ZoteroRibbon.ZoteroRibbonGetSupertip"
        ):
            control_id = button.get("id")
            if not control_id:
                raise ValueError("Localized Ribbon buttons must have an id")
            if control_id in seen_ids:
                raise ValueError(f"Duplicate localized Ribbon control id: {control_id}")
            if "'" in control_id:
                raise ValueError(
                    f"{control_id}: control id must not contain a single quote"
                )
            seen_ids.add(control_id)

            tag = button.get("tag", "")
            if len(tag) > 1024:
                raise ValueError(f"{control_id}: fallback tag exceeds 1024 characters")
            fallback_parts = tag.split("||")
            if len(fallback_parts) != 2:
                raise ValueError(
                    f"{control_id}: tag must contain exactly "
                    "'English label||English supertip'"
                )
            for name, value in zip(("label", "supertip"), fallback_parts):
                if not value.strip():
                    raise ValueError(
                        f"{control_id}: English fallback {name} must be non-empty"
                    )
                if value != value.strip():
                    raise ValueError(
                        f"{control_id}: English fallback {name} has leading/trailing whitespace"
                    )
                validate_xml_text(value, f"{control_id}: English fallback {name}")

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
        data = load_json(path)
        locale = data.get("locale")
        language_ids = data.get("officeLanguageIDs")
        controls = data.get("controls")

        if not isinstance(locale, str) or not locale:
            raise ValueError(f"{path.name}: locale must be a non-empty string")
        if locale != locale.strip():
            raise ValueError(f"{path.name}: locale has leading/trailing whitespace")
        if not LOCALE_RE.fullmatch(locale):
            raise ValueError(
                f"{path.name}: locale may contain only ASCII letters, digits, and hyphens"
            )
        validate_xml_text(locale, f"{path.name}: locale")

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
            if type(language_id) is not int or language_id <= 0:
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
            if set(strings) != {"label", "supertip"}:
                raise ValueError(
                    f"{path.name}: {control_id} must contain exactly label and supertip"
                )
            for key in ("label", "supertip"):
                value = strings.get(key)
                context = f"{path.name}: {control_id}.{key}"
                if not isinstance(value, str) or not value.strip():
                    raise ValueError(f"{context} must be non-empty")
                if value != value.strip():
                    raise ValueError(f"{context} has leading/trailing whitespace")
                validate_xml_text(value, context)

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
        f"{{{DATA_STORE_NS}}}datastoreItem",
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


def validate_localization_xml(data, description):
    root = parse_xml(data, description)
    expected_tag = f"{{{RIBBON_NAMESPACE}}}ribbonLocalization"
    if root.tag != expected_tag:
        raise ValueError(f"{description}: unexpected root element {root.tag!r}")
    if root.get("defaultLocale") != DEFAULT_LOCALE:
        raise ValueError(f"{description}: defaultLocale must be {DEFAULT_LOCALE!r}")
    return root


def validate_item_props(data, description):
    root = parse_xml(data, description)
    expected_tag = f"{{{DATA_STORE_NS}}}datastoreItem"
    if root.tag != expected_tag:
        raise ValueError(f"{description}: unexpected root element {root.tag!r}")
    if root.get(f"{{{DATA_STORE_NS}}}itemID") != DATA_STORE_ITEM_ID:
        raise ValueError(f"{description}: unexpected or missing ds:itemID")

    refs = root.findall(
        f"{{{DATA_STORE_NS}}}schemaRefs/{{{DATA_STORE_NS}}}schemaRef"
    )
    if len(refs) != 1 or refs[0].get(f"{{{DATA_STORE_NS}}}uri") != RIBBON_NAMESPACE:
        raise ValueError(f"{description}: invalid Ribbon localization schemaRef")
    return root


def validate_item_rels(data, description):
    root = parse_xml(data, description)
    if root.tag != f"{{{REL_NS}}}Relationships":
        raise ValueError(f"{description}: unexpected root element {root.tag!r}")

    prop_relationships = [
        rel
        for rel in root.findall(f"{{{REL_NS}}}Relationship")
        if rel.get("Type") == CUSTOM_XML_PROPS_REL_TYPE
    ]
    if (
        len(prop_relationships) != 1
        or prop_relationships[0].get("Target") != "itemProps1.xml"
    ):
        raise ValueError(
            f"{description}: expected exactly one customXmlProps relationship "
            "to itemProps1.xml"
        )
    return root


def validate_generated_xml(localization_xml, item_props, item_rels):
    validate_localization_xml(localization_xml, "generated Ribbon localization XML")
    validate_item_props(item_props, "generated Custom XML properties")
    validate_item_rels(item_rels, "generated Custom XML relationships")


def update_document_rels(data):
    root = parse_xml(data, DOCUMENT_RELS_PART)
    target = "../" + CUSTOM_XML_PART
    relationships = root.findall(f"{{{REL_NS}}}Relationship")

    matches = [
        rel
        for rel in relationships
        if rel.get("Type") == CUSTOM_XML_REL_TYPE and rel.get("Target") == target
    ]
    if len(matches) > 1:
        raise ValueError(
            f"{DOCUMENT_RELS_PART}: duplicate Ribbon localization relationships"
        )
    if len(matches) == 1:
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
    root = parse_xml(data, CONTENT_TYPES_PART)
    part_name = "/" + CUSTOM_XML_PROPS_PART

    xml_defaults = [
        node
        for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
        if node.get("Extension", "").lower() == "xml"
    ]
    if len(xml_defaults) > 1:
        raise ValueError(f"{CONTENT_TYPES_PART}: duplicate xml Default entries")
    if xml_defaults:
        xml_defaults[0].set("ContentType", "application/xml")
    else:
        ET.SubElement(
            root,
            f"{{{CONTENT_TYPES_NS}}}Default",
            {"Extension": "xml", "ContentType": "application/xml"},
        )

    overrides = [
        node
        for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        if node.get("PartName") == part_name
    ]
    if len(overrides) > 1:
        raise ValueError(f"{CONTENT_TYPES_PART}: duplicate override for {part_name}")
    if overrides:
        overrides[0].set("ContentType", CUSTOM_XML_PROPS_CONTENT_TYPE)
    else:
        ET.SubElement(
            root,
            f"{{{CONTENT_TYPES_NS}}}Override",
            {
                "PartName": part_name,
                "ContentType": CUSTOM_XML_PROPS_CONTENT_TYPE,
            },
        )

    return serialize(root)


def check_existing_part_ownership(src, names):
    managed_parts = {
        CUSTOM_XML_PART,
        CUSTOM_XML_PROPS_PART,
        CUSTOM_XML_RELS_PART,
    }
    existing_managed_parts = managed_parts.intersection(names)

    if CUSTOM_XML_PART not in names:
        if existing_managed_parts:
            raise RuntimeError(
                f"{INSTALL_TEMPLATE}: found incomplete custom XML parts at the "
                "paths reserved for Ribbon localization"
            )
        return

    root = parse_xml(src.read(CUSTOM_XML_PART), f"{INSTALL_TEMPLATE}:{CUSTOM_XML_PART}")
    if root.tag != f"{{{RIBBON_NAMESPACE}}}ribbonLocalization":
        raise RuntimeError(
            f"{INSTALL_TEMPLATE}: {CUSTOM_XML_PART} is already used by unrelated "
            "Custom XML data; refusing to overwrite it"
        )


def patch_template(custom_ui, localization_xml, item_props, item_rels):
    if not INSTALL_TEMPLATE.exists():
        raise FileNotFoundError(f"Template not found: {INSTALL_TEMPLATE}")

    with zipfile.ZipFile(INSTALL_TEMPLATE, "r") as src:
        all_names = src.namelist()
        duplicates = [name for name, count in Counter(all_names).items() if count > 1]
        if duplicates:
            raise RuntimeError(
                f"{INSTALL_TEMPLATE}: duplicate ZIP entries: {', '.join(duplicates)}"
            )

        names = set(all_names)
        for required in (
            CUSTOM_UI_PART,
            VBA_PROJECT_PART,
            DOCUMENT_RELS_PART,
            CONTENT_TYPES_PART,
        ):
            if required not in names:
                raise RuntimeError(
                    f"{INSTALL_TEMPLATE}: required package part {required} is missing"
                )

        check_existing_part_ownership(src, names)

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

            with zipfile.ZipFile(temp_path, "r") as built:
                built_names = built.namelist()
                if len(built_names) != len(set(built_names)):
                    raise RuntimeError(
                        f"{temp_path}: duplicate ZIP entries after generation"
                    )
                validate_localization_xml(
                    built.read(CUSTOM_XML_PART),
                    f"{temp_path}:{CUSTOM_XML_PART}",
                )
                validate_item_props(
                    built.read(CUSTOM_XML_PROPS_PART),
                    f"{temp_path}:{CUSTOM_XML_PROPS_PART}",
                )
                validate_item_rels(
                    built.read(CUSTOM_XML_RELS_PART),
                    f"{temp_path}:{CUSTOM_XML_RELS_PART}",
                )

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


def check_xml_file(errors, path, expected, description):
    if not path.exists():
        errors.append(f"missing: {path.relative_to(REPO_ROOT)}")
        return
    actual = path.read_bytes()
    if not xml_semantically_equal(actual, expected):
        errors.append(f"out of date or invalid: {path.relative_to(REPO_ROOT)}")
        return
    try:
        if path == SOURCE_CUSTOM_XML:
            validate_localization_xml(actual, description)
        elif path == SOURCE_CUSTOM_XML_PROPS:
            validate_item_props(actual, description)
        elif path == SOURCE_CUSTOM_XML_RELS:
            validate_item_rels(actual, description)
    except ValueError as exc:
        errors.append(str(exc))


def check_source_vba(errors):
    if not SOURCE_VBA.exists():
        errors.append(f"missing: {SOURCE_VBA.relative_to(REPO_ROOT)}")
        return
    source = SOURCE_VBA.read_text(encoding="utf-8", errors="replace")
    for callback in REQUIRED_VBA_CALLBACKS:
        if callback not in source:
            errors.append(f"{SOURCE_VBA.relative_to(REPO_ROOT)} is missing {callback[:-1]}")


def check(localization_xml, item_props, item_rels):
    expected_sources = {
        SOURCE_CUSTOM_XML: (localization_xml, "source Ribbon localization XML"),
        SOURCE_CUSTOM_XML_PROPS: (item_props, "source Custom XML properties"),
        SOURCE_CUSTOM_XML_RELS: (item_rels, "source Custom XML relationships"),
    }
    errors = []

    for path, (expected, description) in expected_sources.items():
        check_xml_file(errors, path, expected, description)

    check_source_vba(errors)

    if not INSTALL_TEMPLATE.exists():
        errors.append(f"missing: {INSTALL_TEMPLATE.relative_to(REPO_ROOT)}")
        return errors

    try:
        with zipfile.ZipFile(INSTALL_TEMPLATE, "r") as zf:
            all_names = zf.namelist()
            duplicate_names = [
                name for name, count in Counter(all_names).items() if count > 1
            ]
            if duplicate_names:
                errors.append(
                    "install/Zotero.dotm has duplicate ZIP entries: "
                    + ", ".join(duplicate_names)
                )

            names = set(all_names)
            required = (
                CUSTOM_UI_PART,
                VBA_PROJECT_PART,
                CUSTOM_XML_PART,
                CUSTOM_XML_PROPS_PART,
                CUSTOM_XML_RELS_PART,
                DOCUMENT_RELS_PART,
                CONTENT_TYPES_PART,
            )
            for name in required:
                if name not in names:
                    errors.append(f"install/Zotero.dotm is missing {name}")

            if CUSTOM_UI_PART in names:
                if not xml_semantically_equal(
                    zf.read(CUSTOM_UI_PART),
                    SOURCE_CUSTOM_UI.read_bytes(),
                ):
                    errors.append("install/Zotero.dotm has out-of-date customUI.xml")

            if CUSTOM_XML_PART in names:
                actual = zf.read(CUSTOM_XML_PART)
                if not xml_semantically_equal(actual, localization_xml):
                    errors.append(
                        "install/Zotero.dotm has out-of-date Ribbon localization XML"
                    )
                try:
                    validate_localization_xml(
                        actual, "install/Zotero.dotm Ribbon localization XML"
                    )
                except ValueError as exc:
                    errors.append(str(exc))

            if CUSTOM_XML_PROPS_PART in names:
                actual = zf.read(CUSTOM_XML_PROPS_PART)
                if not xml_semantically_equal(actual, item_props):
                    errors.append(
                        "install/Zotero.dotm has out-of-date Custom XML properties"
                    )
                try:
                    validate_item_props(
                        actual, "install/Zotero.dotm Custom XML properties"
                    )
                except ValueError as exc:
                    errors.append(str(exc))

            if CUSTOM_XML_RELS_PART in names:
                actual = zf.read(CUSTOM_XML_RELS_PART)
                if not xml_semantically_equal(actual, item_rels):
                    errors.append(
                        "install/Zotero.dotm has out-of-date Custom XML relationships"
                    )
                try:
                    validate_item_rels(
                        actual, "install/Zotero.dotm Custom XML relationships"
                    )
                except ValueError as exc:
                    errors.append(str(exc))

            if DOCUMENT_RELS_PART in names:
                try:
                    rels = parse_xml(
                        zf.read(DOCUMENT_RELS_PART),
                        "install/Zotero.dotm document relationships",
                    )
                    target = "../" + CUSTOM_XML_PART
                    matches = [
                        rel
                        for rel in rels.findall(f"{{{REL_NS}}}Relationship")
                        if rel.get("Type") == CUSTOM_XML_REL_TYPE
                        and rel.get("Target") == target
                    ]
                    if len(matches) != 1:
                        errors.append(
                            "install/Zotero.dotm must contain exactly one "
                            "Ribbon localization document relationship"
                        )
                    if CUSTOM_XML_PART not in names:
                        errors.append(
                            "install/Zotero.dotm Ribbon localization relationship "
                            "target is missing"
                        )
                except ValueError as exc:
                    errors.append(str(exc))

            if CONTENT_TYPES_PART in names:
                try:
                    types = parse_xml(
                        zf.read(CONTENT_TYPES_PART),
                        "install/Zotero.dotm content types",
                    )
                    xml_defaults = [
                        node
                        for node in types.findall(f"{{{CONTENT_TYPES_NS}}}Default")
                        if node.get("Extension", "").lower() == "xml"
                        and node.get("ContentType") == "application/xml"
                    ]
                    if len(xml_defaults) != 1:
                        errors.append(
                            "install/Zotero.dotm must map the xml extension "
                            "to application/xml exactly once"
                        )

                    part_name = "/" + CUSTOM_XML_PROPS_PART
                    overrides = [
                        node
                        for node in types.findall(f"{{{CONTENT_TYPES_NS}}}Override")
                        if node.get("PartName") == part_name
                        and node.get("ContentType") == CUSTOM_XML_PROPS_CONTENT_TYPE
                    ]
                    if len(overrides) != 1:
                        errors.append(
                            "install/Zotero.dotm is missing or duplicates the "
                            "Custom XML properties content type"
                        )
                except ValueError as exc:
                    errors.append(str(exc))

            if CUSTOM_XML_RELS_PART in names:
                try:
                    rels = validate_item_rels(
                        zf.read(CUSTOM_XML_RELS_PART),
                        "install/Zotero.dotm Custom XML relationships",
                    )
                    targets = [
                        rel.get("Target")
                        for rel in rels.findall(f"{{{REL_NS}}}Relationship")
                        if rel.get("Type") == CUSTOM_XML_PROPS_REL_TYPE
                    ]
                    for target in targets:
                        target_path = "customXml/" + target
                        if target_path not in names:
                            errors.append(
                                "install/Zotero.dotm Custom XML properties "
                                f"relationship target is missing: {target_path}"
                            )
                except ValueError:
                    pass

    except (zipfile.BadZipFile, OSError) as exc:
        errors.append(f"install/Zotero.dotm cannot be read as an OOXML package: {exc}")

    return errors


def main():
    parser = argparse.ArgumentParser(
        description="Build locale files into the Zotero Word Ribbon localization store."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help=(
            "Check generated localization sources and package metadata without "
            "modifying files. This does not compile or inspect VBA bytecode."
        ),
    )
    args = parser.parse_args()

    control_ids = get_ribbon_controls()
    locales = load_locales(control_ids)
    localization_xml = build_localization_xml(locales, control_ids)
    item_props = build_item_props()
    item_rels = build_item_rels()
    validate_generated_xml(localization_xml, item_props, item_rels)
    custom_ui = SOURCE_CUSTOM_UI.read_bytes()

    if args.check:
        errors = check(localization_xml, item_props, item_rels)
        if errors:
            for error in errors:
                print(error)
            raise SystemExit(1)
        print("Ribbon localization data and package metadata are up to date.")
        print("Compiled VBA bytecode is not verified by this check.")
        return

    write_sources(localization_xml, item_props, item_rels)
    patch_template(custom_ui, localization_xml, item_props, item_rels)
    print(f"Built {len(locales)} Ribbon locale(s) into install/Zotero.dotm.")
    print(
        "Note: this updates Ribbon/package localization data only; "
        "compile VBA in Word whenever ZoteroRibbon.bas changes."
    )


if __name__ == "__main__":
    main()
