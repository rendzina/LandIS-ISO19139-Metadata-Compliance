#!/usr/bin/env python3
"""
Batch extraction of metadata from ISO 19139 / ArcGIS XML files into a single Excel workbook.

The input folder is a script parameter (default: 'xml'). The output file is named
metadata_export_<foldername>.xlsx so it reflects the source folder. The workbook contains:
  - A "Metadata Export" sheet: one row per file, one column per attribute, with a second
    row labelling each attribute as mandatory, optional, or conditional (ISO 19139/INSPIRE).
  - A "Compliance Summary" sheet: per-file ISO 19139 compliance (Yes/No) and list of
    missing mandatory fields.

Designed for Esri/ArcGIS ISO 19139-style metadata (e.g. from ArcGIS Online) and aligned
with INSPIRE Regulation 1205/2008 for mandatory/optional/conditional classification.
"""

import argparse
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import html
import re
from pathlib import Path
from collections import OrderedDict

# ISO 19139 / INSPIRE obligation per exported field name (Regulation 1205/2008, INSPIRE TG).
# Used for: (1) the optionality row in "Metadata Export", (2) which fields count as
# mandatory in "Compliance Summary". Keys must match the display names used in
# extract_all_fields(). Unknown fields default to 'optional' in get_field_obligation().
FIELD_OBLIGATION = {
    "Resource Title": "mandatory",
    "Abstract": "mandatory",
    "Topic Category": "mandatory",
    "Keywords": "mandatory",
    "Geographic West Bounding Longitude": "mandatory",
    "Geographic East Bounding Longitude": "mandatory",
    "Geographic North Bounding Latitude": "mandatory",
    "Geographic South Bounding Latitude": "mandatory",
    "Data Language": "mandatory",
    "Scale Denominator": "mandatory",
    "Contact Organisation Name": "mandatory",
    "Contact Email Address": "mandatory",
    "Contact Role": "mandatory",
    "Distribution Online Resource Linkage": "mandatory",
    "Lineage Statement": "mandatory",
    "Data Quality Scope Level": "mandatory",
    "Metadata Language Code": "mandatory",
    "Metadata Date Stamp": "mandatory",
    "Metadata Scope Code": "mandatory",
    "Access Constraints": "mandatory",
    "Conformance Specification Title": "mandatory",
    "Conformance Pass": "mandatory",
    "Publication Date": "conditional",
    "Reference System Code": "conditional",
    "Reference System Code Space": "conditional",
    "Other Constraints": "mandatory",
    "Use Limitation": "mandatory",
}
# All other fields (ArcGIS-specific, optional contact details, etc.) are optional unless listed above.


def get_field_obligation(field_name):
    """
    Return the ISO 19139/INSPIRE obligation for an exported field name.

    Used to populate the optionality row and to determine mandatory fields for
    the compliance summary. Any field not in FIELD_OBLIGATION (including
    "Other Keywords (...)" variants) is treated as optional.

    Args:
        field_name: The display name of the attribute as used in the export (e.g.
                    "Resource Title", "Abstract").

    Returns:
        One of "mandatory", "optional", or "conditional".
    """
    if field_name in FIELD_OBLIGATION:
        return FIELD_OBLIGATION[field_name]
    if field_name.startswith("Other Keywords ("):
        return "optional"
    return "optional"


def clean_text(text):
    """
    Normalise text by decoding HTML entities, stripping tags, and collapsing whitespace.

    Args:
        text: Raw string possibly containing HTML (e.g. from idAbs or useLimit).
              Can be None.

    Returns:
        Cleaned non-None string; empty string if input is None or empty after cleaning.
    """
    if text is None:
        return ""
    # Decode HTML entities
    text = html.unescape(text)
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    # Clean up whitespace
    text = ' '.join(text.split())
    return text.strip()


def get_text(element, default=""):
    """
    Extract and clean all text content from an XML element, including nested children.

    Handles elements that contain both direct text and child elements (e.g. div with
    spans). Result is passed through clean_text() to remove HTML and normalise space.

    Args:
        element: An xml.etree.ElementTree.Element, or None.
        default: Value to return if element is None or has no text content.

    Returns:
        Cleaned string of concatenated text; default if element is None or empty.
    """
    if element is None:
        return default
    if element.text:
        return clean_text(element.text)
    # Get all text content including nested elements
    text_parts = []
    if element.text:
        text_parts.append(element.text.strip())
    for child in element:
        if child.text:
            text_parts.append(child.text.strip())
        if child.tail:
            text_parts.append(child.tail.strip())
    return clean_text(' '.join(text_parts))


def get_attribute_value(element, attr_name, default=""):
    """
    Return the value of an XML attribute, or a default if missing or element is None.

    Args:
        element: An xml.etree.ElementTree.Element, or None.
        attr_name: Attribute name (e.g. 'value', 'codeListValue').
        default: Value to return when element is None or attribute is absent.

    Returns:
        Attribute value string, or default.
    """
    if element is None:
        return default
    return element.get(attr_name, default)


def extract_all_fields(root):
    """
    Extract all supported metadata fields from an ISO 19139 / ArcGIS metadata XML root.

    Walks Esri, dataIdInfo, mdContact, eainfo, spatRepInfo, refSysInfo, dqInfo,
    distInfo, mdMaint, mdLang, mdHrLv, spdoinfo, and related elements. Only non-empty
    values are stored. If the same field is set more than once (e.g. multiple keywords),
    values are concatenated with " | ".

    Args:
        root: The root Element of the parsed XML (e.g. from ET.parse(path).getroot()).
              Can be Esri/ArcGIS-style or standard ISO 19139.

    Returns:
        OrderedDict mapping display field names (e.g. "Resource Title") to string values.
        Order of insertion is preserved for consistent column ordering when combined
        with other files.
    """
    fields = OrderedDict()

    def add_field(field_name, value):
        """Add a non-empty field to the accumulator. If key exists, append with ' | '."""
        if value:
            # If field already exists, append with separator
            if field_name in fields:
                if fields[field_name]:
                    fields[field_name] = f"{fields[field_name]} | {value}"
                else:
                    fields[field_name] = value
            else:
                fields[field_name] = str(value)
    
    # Extract Esri metadata
    esri = root.find('.//Esri')
    if esri is not None:
        add_field("ArcGIS Format", get_text(esri.find('ArcGISFormat')))
        add_field("ArcGIS Profile", get_text(esri.find('ArcGISProfile')))
        add_field("Creation Date", get_text(esri.find('CreaDate')))
        add_field("Creation Time", get_text(esri.find('CreaTime')))
        add_field("Modification Date", get_text(esri.find('ModDate')))
        add_field("Modification Time", get_text(esri.find('ModTime')))
        
        # Data Properties
        data_props = esri.find('.//DataProperties')
        if data_props is not None:
            item_props = data_props.find('.//itemProps')
            if item_props is not None:
                add_field("Item Name", get_text(item_props.find('itemName')))
                add_field("Content Type", get_text(item_props.find('imsContentType')))
                
                # Native extent
                native_ext = item_props.find('.//nativeExtBox')
                if native_ext is not None:
                    add_field("West Bounding Longitude", get_text(native_ext.find('westBL')))
                    add_field("East Bounding Longitude", get_text(native_ext.find('eastBL')))
                    add_field("South Bounding Latitude", get_text(native_ext.find('southBL')))
                    add_field("North Bounding Latitude", get_text(native_ext.find('northBL')))
                
                # Portal details
                portal = item_props.find('.//portalDetails')
                if portal is not None:
                    add_field("Thumbnail URL", get_text(portal.find('thumbnailURL')))
            
            # Coordinate reference
            coord_ref = data_props.find('.//coordRef')
            if coord_ref is not None:
                add_field("Coordinate System Type", get_text(coord_ref.find('type')))
                add_field("Geographic CS Name", get_text(coord_ref.find('geogcsn')))
                add_field("Projected CS Name", get_text(coord_ref.find('projcsn')))
                add_field("Coordinate System Units", get_text(coord_ref.find('csUnits')))
        
        # Scale range
        scale_range = esri.find('.//scaleRange')
        if scale_range is not None:
            add_field("Minimum Scale", get_text(scale_range.find('minScale')))
            add_field("Maximum Scale", get_text(scale_range.find('maxScale')))
    
    # Extract Data Identification Info
    data_id = root.find('.//dataIdInfo')
    if data_id is not None:
        # Abstract
        abstract = data_id.find('idAbs')
        if abstract is not None:
            add_field("Abstract", get_text(abstract))
        
        # Citation
        citation = data_id.find('idCitation')
        if citation is not None:
            add_field("Resource Title", get_text(citation.find('resTitle')))
            add_field("Resource Alternative Title", get_text(citation.find('resAltTitle')))
            add_field("Collection Title", get_text(citation.find('collTitle')))
            
            date_elem = citation.find('.//date/pubDate')
            if date_elem is not None:
                add_field("Publication Date", get_text(date_elem))
            
            pres_form = citation.find('.//presForm/PresFormCd')
            if pres_form is not None:
                add_field("Presentation Form", get_attribute_value(pres_form, 'value'))
        
        # Extent
        data_ext = data_id.find('dataExt')
        if data_ext is not None:
            add_field("Extent Description", get_text(data_ext.find('exDesc')))
            
            geo_bbox = data_ext.find('.//GeoBndBox')
            if geo_bbox is not None:
                add_field("Geographic West Bounding Longitude", get_text(geo_bbox.find('westBL')))
                add_field("Geographic East Bounding Longitude", get_text(geo_bbox.find('eastBL')))
                add_field("Geographic North Bounding Latitude", get_text(geo_bbox.find('northBL')))
                add_field("Geographic South Bounding Latitude", get_text(geo_bbox.find('southBL')))
        
        # Keywords
        keywords = data_id.findall('.//searchKeys/keyword')
        keyword_list = [get_text(kw) for kw in keywords if get_text(kw)]
        if keyword_list:
            add_field("Keywords", ', '.join(keyword_list))
        
        # Purpose
        add_field("Purpose", get_text(data_id.find('idPurp')))
        
        # Credit
        add_field("Credit", get_text(data_id.find('idCredit')))
        
        # Constraints
        use_limit = data_id.find('.//useLimit')
        if use_limit is not None:
            add_field("Use Limitation", get_text(use_limit))
        
        access_const = data_id.find('.//accessConsts/RestrictCd')
        if access_const is not None:
            add_field("Access Constraints", get_attribute_value(access_const, 'value'))
        
        other_const = data_id.find('.//othConsts')
        if other_const is not None:
            add_field("Other Constraints", get_text(other_const))
        
        # Language
        lang_code = data_id.find('.//dataLang/languageCode')
        if lang_code is not None:
            add_field("Data Language", get_attribute_value(lang_code, 'value'))
        
        country_code = data_id.find('.//dataLang/countryCode')
        if country_code is not None:
            add_field("Data Country Code", get_attribute_value(country_code, 'value'))
        
        # Character Set
        char_set = data_id.find('.//dataChar/CharSetCd')
        if char_set is not None:
            add_field("Character Set", get_attribute_value(char_set, 'value'))
        
        # Spatial Representation Type
        spat_rep = data_id.find('.//spatRpType/SpatRepTypCd')
        if spat_rep is not None:
            add_field("Spatial Representation Type", get_attribute_value(spat_rep, 'value'))
        
        # Scale
        scale = data_id.find('.//dataScale/equScale/rfDenom')
        if scale is not None:
            add_field("Scale Denominator", get_text(scale))
        
        # Environment
        envir = data_id.find('envirDesc')
        if envir is not None:
            add_field("Environment Description", get_text(envir))
        
        # Status
        status = data_id.find('.//idStatus/ProgCd')
        if status is not None:
            add_field("Status", get_attribute_value(status, 'value'))
        
        # Graphic Overview
        graph_over = data_id.find('graphOver')
        if graph_over is not None:
            add_field("Graphic File Name", get_text(graph_over.find('bgFileName')))
            add_field("Graphic File Description", get_text(graph_over.find('bgFileDesc')))
            add_field("Graphic File Type", get_text(graph_over.find('bgFileType')))
        
        # Maintenance
        maint = data_id.find('.//resMaint/maintFreq/MaintFreqCd')
        if maint is not None:
            add_field("Maintenance Frequency", get_attribute_value(maint, 'value'))
        
        # Topic Category
        topic_cat = data_id.find('.//tpCat/TopicCatCd')
        if topic_cat is not None:
            add_field("Topic Category", get_attribute_value(topic_cat, 'value'))
        
        # Other Keywords
        other_keys = data_id.findall('.//otherKeys')
        for other_key in other_keys:
            thesa_name = get_text(other_key.find('.//thesaName/resTitle'))
            keywords = [get_text(kw) for kw in other_key.findall('keyword') if get_text(kw)]
            if keywords:
                key_name = f"Other Keywords ({thesa_name})" if thesa_name else "Other Keywords"
                add_field(key_name, ', '.join(keywords))
    
    # Extract Contact Information
    contact = root.find('.//mdContact')
    if contact is not None:
        add_field("Contact Individual Name", get_text(contact.find('rpIndName')))
        add_field("Contact Organisation Name", get_text(contact.find('rpOrgName')))
        add_field("Contact Position Name", get_text(contact.find('rpPosName')))
        
        cnt_info = contact.find('.//rpCntInfo')
        if cnt_info is not None:
            # Address
            address = cnt_info.find('.//cntAddress')
            if address is not None:
                add_field("Contact Email Address", get_text(address.find('eMailAdd')))
                add_field("Contact Delivery Point", get_text(address.find('delPoint')))
                add_field("Contact City", get_text(address.find('city')))
                add_field("Contact Administrative Area", get_text(address.find('adminArea')))
                add_field("Contact Postal Code", get_text(address.find('postCode')))
                add_field("Contact Country", get_text(address.find('country')))
            
            # Phone
            phone = cnt_info.find('.//cntPhone/voiceNum')
            if phone is not None:
                add_field("Contact Phone Number", get_text(phone))
            
            # Online Resource
            online = cnt_info.find('.//cntOnlineRes/linkage')
            if online is not None:
                add_field("Contact Online Resource", get_text(online))
            
            # Hours
            hours = cnt_info.find('.//cntHours')
            if hours is not None:
                add_field("Contact Hours", get_text(hours))
            
            # Instructions
            instr = cnt_info.find('.//cntInstr')
            if instr is not None:
                add_field("Contact Instructions", get_text(instr))
        
        role = contact.find('.//role/RoleCd')
        if role is not None:
            add_field("Contact Role", get_attribute_value(role, 'value'))
    
    # Extract Attribute Definitions (eainfo)
    eainfo = root.find('.//eainfo/detailed')
    if eainfo is not None:
        enttyp = eainfo.find('enttyp')
        if enttyp is not None:
            add_field("Entity Type Label", get_text(enttyp.find('enttypl')))
            add_field("Entity Type Type", get_text(enttyp.find('enttypt')))
            add_field("Entity Type Count", get_text(enttyp.find('enttypc')))
        
        # Process all attributes - store as a summary
        attributes = eainfo.findall('.//attr')
        attr_summaries = []
        for attr in attributes:
            attr_label = get_text(attr.find('attrlabl'))
            if attr_label:
                attr_summaries.append(attr_label)
        if attr_summaries:
            add_field("Attribute Names", ', '.join(attr_summaries))
    
    # Extract Spatial Representation Info
    spat_rep_info = root.find('.//spatRepInfo')
    if spat_rep_info is not None:
        top_level = spat_rep_info.find('.//topLvl/TopoLevCd')
        if top_level is not None:
            add_field("Topology Level", get_attribute_value(top_level, 'value'))
        
        geo_objs = spat_rep_info.find('.//geometObjs')
        if geo_objs is not None:
            geo_type = geo_objs.find('.//geoObjTyp/GeoObjTypCd')
            if geo_type is not None:
                add_field("Geometry Object Type", get_attribute_value(geo_type, 'value'))
            
            geo_count = geo_objs.find('.//geoObjCnt')
            if geo_count is not None:
                add_field("Geometry Object Count", get_text(geo_count))
    
    # Extract Reference System Info
    ref_sys = root.find('.//refSysInfo/RefSystem/refSysID')
    if ref_sys is not None:
        ident_code = ref_sys.find('identCode')
        if ident_code is not None:
            add_field("Reference System Code", get_attribute_value(ident_code, 'code'))
        
        code_space = ref_sys.find('idCodeSpace')
        if code_space is not None:
            add_field("Reference System Code Space", get_text(code_space))
        
        version = ref_sys.find('idVersion')
        if version is not None:
            add_field("Reference System Version", get_text(version))
    
    # Extract Data Quality Info
    dq_info = root.find('.//dqInfo')
    if dq_info is not None:
        scope = dq_info.find('.//scpLvl/ScopeCd')
        if scope is not None:
            add_field("Data Quality Scope Level", get_attribute_value(scope, 'value'))
        
        lineage = dq_info.find('.//dataLineage/statement')
        if lineage is not None:
            add_field("Lineage Statement", get_text(lineage))
        
        report = dq_info.find('.//report')
        if report is not None:
            report_type = report.get('type', '')
            add_field("Quality Report Type", report_type)
            
            con_spec = report.find('.//conSpec/resTitle')
            if con_spec is not None:
                add_field("Conformance Specification Title", get_text(con_spec))
            
            con_pass = report.find('.//conPass')
            if con_pass is not None:
                add_field("Conformance Pass", get_text(con_pass))
    
    # Extract Distribution Info
    dist_info = root.find('.//distInfo')
    if dist_info is not None:
        online_src = dist_info.find('.//onLineSrc')
        if online_src is not None:
            linkage = online_src.find('linkage')
            if linkage is not None:
                add_field("Distribution Online Resource Linkage", get_text(linkage))
            
            protocol = online_src.find('protocol')
            if protocol is not None:
                add_field("Distribution Protocol", get_text(protocol))
            
            or_name = online_src.find('orName')
            if or_name is not None:
                add_field("Distribution Online Resource Name", get_text(or_name))
            
            or_desc = online_src.find('orDesc')
            if or_desc is not None:
                add_field("Distribution Online Resource Description", get_text(or_desc))
    
    # Extract Maintenance Info
    md_maint = root.find('.//mdMaint')
    if md_maint is not None:
        maint_freq = md_maint.find('.//maintFreq/MaintFreqCd')
        if maint_freq is not None:
            add_field("Metadata Maintenance Frequency", get_attribute_value(maint_freq, 'value'))
    
    # Extract Metadata Language
    md_lang = root.find('.//mdLang')
    if md_lang is not None:
        lang_code = md_lang.find('languageCode')
        if lang_code is not None:
            add_field("Metadata Language Code", get_attribute_value(lang_code, 'value'))
        
        country_code = md_lang.find('countryCode')
        if country_code is not None:
            add_field("Metadata Country Code", get_attribute_value(country_code, 'value'))
    
    # Extract Metadata Hierarchy Level
    md_hr_lv = root.find('.//mdHrLv')
    if md_hr_lv is not None:
        scope_cd = md_hr_lv.find('ScopeCd')
        if scope_cd is not None:
            add_field("Metadata Scope Code", get_attribute_value(scope_cd, 'value'))
    
    hr_lv_name = root.find('.//mdHrLvName')
    if hr_lv_name is not None:
        add_field("Metadata Hierarchy Level Name", get_text(hr_lv_name))
    
    # Extract Spatial Domain Info
    spdo_info = root.find('.//spdoinfo')
    if spdo_info is not None:
        esri_term = spdo_info.find('.//esriterm')
        if esri_term is not None:
            name = esri_term.get('Name', '')
            add_field("Feature Name", name)
            
            feat_type = esri_term.find('efeatyp')
            if feat_type is not None:
                add_field("Feature Type", get_text(feat_type))
            
            feat_geom = esri_term.find('efeageom')
            if feat_geom is not None:
                add_field("Feature Geometry Code", get_attribute_value(feat_geom, 'code'))
    
    # Extract Metadata Standard
    md_std_name = root.find('.//mdStanName')
    if md_std_name is not None:
        add_field("Metadata Standard Name", get_text(md_std_name))
    
    md_std_ver = root.find('.//mdStanVer')
    if md_std_ver is not None:
        add_field("Metadata Standard Version", get_text(md_std_ver))
    
    # Extract File ID
    md_file_id = root.find('.//mdFileID')
    if md_file_id is not None:
        add_field("Metadata File ID", get_text(md_file_id))
    
    md_char = root.find('.//mdChar')
    if md_char is not None:
        char_set = md_char.find('CharSetCd')
        if char_set is not None:
            add_field("Metadata Character Set", get_attribute_value(char_set, 'value'))
    
    md_date = root.find('.//mdDateSt')
    if md_date is not None:
        add_field("Metadata Date Stamp", get_text(md_date))
    
    return fields


def process_all_xml_files(xml_folder):
    """
    Discover and process every .xml file in the given folder.

    Each file is parsed and passed to extract_all_fields(). Failures for a single
    file are logged but do not stop the run. The union of all attribute names
    across successful files is returned as the canonical column set.

    Args:
        xml_folder: Path to the directory containing .xml metadata files (str or Path).

    Returns:
        Tuple (all_data, sorted_field_names), or (None, None) if the folder does not
        exist or contains no .xml files.
        - all_data: Dict mapping filename (str) to OrderedDict of field name -> value.
        - sorted_field_names: Sorted list of all unique attribute names found.
    """
    xml_folder = Path(xml_folder)
    xml_files = sorted(xml_folder.glob("*.xml"))
    
    if not xml_files:
        print(f"No XML files found in {xml_folder}")
        return None, None
    
    print(f"Found {len(xml_files)} XML files to process")
    
    all_data = {}  # filename -> fields dictionary
    all_field_names = set()  # Collect all unique field names
    
    for xml_file in xml_files:
        filename = xml_file.name
        print(f"Processing: {filename}")
        
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            fields = extract_all_fields(root)
            
            all_data[filename] = fields
            all_field_names.update(fields.keys())
            
        except ET.ParseError as e:
            print(f"  Error parsing {filename}: {e}")
        except Exception as e:
            print(f"  Error processing {filename}: {e}")
    
    # Sort field names for consistent column order
    sorted_field_names = sorted(all_field_names)
    
    return all_data, sorted_field_names


def compute_compliance(all_data, field_names):
    """
    Assess ISO 19139 compliance for each file based on presence of mandatory fields.

    A file is considered compliant only if every field marked as "mandatory" in
    FIELD_OBLIGATION is present in that file's extracted data and has a non-empty
    value. Conditional and optional fields do not affect the compliant flag.

    Args:
        all_data: Dict mapping filename to dict of field name -> value (as from
                  process_all_xml_files).
        field_names: List of all attribute names (column set).

    Returns:
        List of dicts, one per file (sorted by filename), each with keys:
        - "Filename": str
        - "Compliant": "Yes" or "No"
        - "Missing mandatory": comma-separated list of missing mandatory field names
        - "Missing count": int
    """
    mandatory_fields = [fn for fn in field_names if get_field_obligation(fn) == "mandatory"]
    results = []
    for filename in sorted(all_data.keys()):
        fields = all_data[filename]
        missing = [fn for fn in mandatory_fields if not (fields.get(fn) or "").strip()]
        compliant = len(missing) == 0
        results.append({
            "Filename": filename,
            "Compliant": "Yes" if compliant else "No",
            "Missing mandatory": ", ".join(missing) if missing else "",
            "Missing count": len(missing),
        })
    return results


def create_excel_matrix(all_data, field_names, output_file):
    """
    Build the Excel workbook and write it to disk.

    Creates two sheets:
    1. "Compliance Summary" (first): one row per file with Filename, ISO 19139
       compliant (Yes/No), Missing mandatory fields, Missing count.
    2. "Metadata Export": Row 1 = headers (Filename + attribute names), Row 2 =
       optionality (mandatory/optional/conditional per column), Row 3+ = one row per
       file with filename and attribute values. Freezes header and optionality rows
       and the filename column. Applies styling and text wrapping.

    Args:
        all_data: Dict mapping filename to dict of field name -> value.
        field_names: List of attribute names defining column order.
        output_file: Path (str or Path) for the output .xlsx file.

    Returns:
        None. Writes the file and prints a short summary to stdout.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Metadata Export"

    # Row 1: Header row – Filename + all field names
    headers = ['Filename'] + field_names
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Row 2: Optionality row – blank for Filename, then mandatory/optional/conditional per column
    ws.cell(row=2, column=1).value = ""  # Filename column
    obligation_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    obligation_font = Font(italic=True)
    for col_num, field_name in enumerate(field_names, 2):
        cell = ws.cell(row=2, column=col_num)
        cell.value = get_field_obligation(field_name)
        cell.fill = obligation_fill
        cell.font = obligation_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Data rows (starting at row 3)
    for row_idx, filename in enumerate(sorted(all_data.keys()), start=3):
        fields = all_data[filename]
        ws.cell(row=row_idx, column=1).value = filename
        for col_num, field_name in enumerate(field_names, 2):
            val = fields.get(field_name, '')
            ws.cell(row=row_idx, column=col_num).value = val

    # Auto-adjust column widths
    for col_num, header in enumerate(headers, 1):
        max_length = len(str(header))
        column_letter = get_column_letter(col_num)
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(row=r, column=col_num)
            if cell.value and len(str(cell.value)) > max_length:
                max_length = min(len(str(cell.value)), 100)
        ws.column_dimensions[column_letter].width = min(max_length + 2, 100)

    # Freeze header and optionality row plus filename column
    ws.freeze_panes = 'B3'

    # Text wrapping for data cells
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')

    # Compliance Summary sheet
    compliance = compute_compliance(all_data, field_names)
    ws_summary = wb.create_sheet("Compliance Summary", 0)
    summary_headers = ["Filename", "ISO 19139 compliant", "Missing mandatory fields", "Missing count"]
    for col_num, h in enumerate(summary_headers, 1):
        c = ws_summary.cell(row=1, column=col_num)
        c.value = h
        c.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row_idx, rec in enumerate(compliance, start=2):
        ws_summary.cell(row=row_idx, column=1).value = rec["Filename"]
        ws_summary.cell(row=row_idx, column=2).value = rec["Compliant"]
        ws_summary.cell(row=row_idx, column=3).value = rec["Missing mandatory"]
        ws_summary.cell(row=row_idx, column=4).value = rec["Missing count"]
    for col_num in range(1, 5):
        ws_summary.column_dimensions[get_column_letter(col_num)].width = 24
    for row in ws_summary.iter_rows(min_row=2, max_row=ws_summary.max_row, min_col=3, max_col=3):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')

    wb.save(output_file)
    print(f"\nExcel file created successfully: {output_file}")
    print(f"Total files processed: {len(all_data)}")
    print(f"Total unique attributes: {len(field_names)}")
    compliant_count = sum(1 for r in compliance if r["Compliant"] == "Yes")
    print(f"ISO 19139 compliance: {compliant_count} compliant, {len(compliance) - compliant_count} with missing mandatory fields")


def parse_args():
    """
    Parse command-line arguments for input folder and derive output filename.

    Returns:
        argparse.Namespace with 'input_folder' (Path) and 'output_file' (str).
        output_file is of the form metadata_export_<foldername>.xlsx so the
        Excel file name reflects the source folder.
    """
    parser = argparse.ArgumentParser(
        description="Extract metadata from ISO 19139 / ArcGIS XML files into an Excel workbook."
    )
    parser.add_argument(
        "input_folder",
        nargs="?",
        default="xml",
        help="Folder containing .xml metadata files (default: xml)",
    )
    args = parser.parse_args()
    xml_folder = Path(args.input_folder)
    folder_name = xml_folder.name
    output_file = f"metadata_export_{folder_name}.xlsx"
    args.input_folder = xml_folder
    args.output_file = output_file
    return args


def main():
    """
    Entry point: run batch extraction and write an Excel workbook.

    The input folder is a script parameter (default: 'xml'). The output file is
    named metadata_export_<foldername>.xlsx so it reflects the source folder.
    Exits with an error if the folder is missing. On success, prints progress
    per file and final counts (files processed, unique attributes, compliant
    vs non-compliant).
    """
    args = parse_args()
    xml_folder = args.input_folder
    output_file = args.output_file

    if not xml_folder.exists():
        print(f"Error: XML folder not found at {xml_folder}")
        return

    print(f"Processing XML files from: {xml_folder}")
    print("-" * 60)

    try:
        all_data, field_names = process_all_xml_files(xml_folder)

        if all_data is None:
            return

        print("-" * 60)
        print(f"Creating Excel file: {output_file}")
        create_excel_matrix(all_data, field_names, output_file)

        print("\nExtraction complete!")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
