# LandIS : ISO 19139 Compliance — XML Metadata Extractor and Reporter

Batch extraction of metadata from ISO 19139 / ArcGIS XML metadata files into a single Excel workbook, with an optionality row and an ISO 19139 compliance summary per file.

Tool background: ESRI ArcGIS Online (AGOL) can hold and export metadata files for each of the feature layers it holds. Metadata can be edited within AGOL, or imported in from another source, e.g. ArcGIS Pro. XML files can be output against all the feature layers as xml files. This tool will process the files and provide a summary statement of the compliance of the metadata with the ISO19139 standard. The tool also outputs all the metadata from each of the XML files processed into one spreadsheet output for ease of access. The tool can therefore be used as part of a workflow to successively check for compliance of metadata with the ISO19139 standard.

The code has been provided as part of the open source output for the LandIS Portal database conversion project. The tool should work with any XML files.

http://www.landis.org.uk

Author: Stephen Hallett, Cranfield University
Date: 04-02-2026

---

## Overview

This project processes **all** `.xml` metadata files in a given folder (default: `xml`) and produces **one Excel workbook** whose name includes the folder name (e.g. `metadata_export_xml.xlsx`) containing:

1. **Metadata Export** sheet  
   - One **row per XML file**, one **column per metadata attribute**.  
   - **Row 1:** Column headers (Filename + attribute names).  
   - **Row 2:** Optionality row — each attribute column is labelled **mandatory**, **optional**, or **conditional** (ISO 19139 / INSPIRE).  
   - **Row 3 onward:** Extracted values; empty cells mean the attribute was absent in that file.

2. **Compliance Summary** sheet  
   - One row per XML file with:  
     - **Filename**  
     - **ISO 19139 compliant** (Yes/No)  
     - **Missing mandatory fields** (comma-separated list)  
     - **Missing count**

Compliance is based on a fixed set of **mandatory** fields derived from INSPIRE / Regulation 1205/2008. A file is “compliant” only if every such mandatory field is present and non-empty in the extracted data.

---

## Features

- **Batch processing:** Discovers and processes every `.xml` file in the configured folder.
- **Unified attribute set:** Collects all unique attribute names across files; each file gets a row with values or blanks.
- **ISO 19139 optionality:** Second row on the export sheet shows whether each column is **mandatory**, **optional**, or **conditional**.
- **Compliance report:** Per-file summary of ISO 19139 compliance and list of missing mandatory fields.
- **Rich extraction:** Supports both Esri/ArcGIS-style and ISO 19139-style elements (data identification, contact, extent, keywords, distribution, data quality, reference system, etc.).
- **Robust text handling:** Strips HTML tags and decodes HTML entities in text values.
- **Error resilience:** A single failing file does not stop the run; errors are reported and the rest are still processed.

---

## Requirements

- **Python:** 3.6 or newer  
- **Dependencies:**  
  - `openpyxl` — Excel (.xlsx) read/write  
  - `lxml` (optional) — faster XML parsing; standard library `xml.etree.ElementTree` is used if `lxml` is not installed  

---

## Installation

```bash
# Clone or copy the project, then from the project root:
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install openpyxl lxml
```

---

## Usage

From the project root, with the virtual environment activated:

```bash
# Use the default folder 'xml'
python extract_metadata.py

# Use a specific folder (output will be metadata_export_<foldername>.xlsx)
python extract_metadata.py XML_Folder

# Use a subfolder path (foldername is the last component, e.g. 'Public')
python extract_metadata.py XML_Exports/Public
```

Without activating the venv (macOS/Linux):

```bash
./venv/bin/python extract_metadata.py XML_Exports/Public
```

- **Input:** All `.xml` files in the given folder (default: `xml`).  
- **Output:** `metadata_export_<foldername>.xlsx` in the current directory (e.g. `metadata_export_xml.xlsx` for the default folder, `metadata_export_Public.xlsx` for `XML_Exports/Public`).  
- **Console:** Progress lines per file plus a final summary (file count, attribute count, compliant vs non-compliant count).

### Conformance checker (strict namespace check)

A second script, `check_conformance.py`, performs a **strict namespace-aware** ISO 19139 conformance check. It only recognises XML whose root is `gmd:MD_Metadata` (official ISO 19139 namespaces). Use it for properly namespaced ISO 19139 metadata; ArcGIS-style exports without these namespaces are skipped and listed in an Errors sheet.

```bash
# Default folder 'xml'
python check_conformance.py

# Specific folder (output: conformance_report_<foldername>.xlsx)
python check_conformance.py V4/Public_ISO

./venv/bin/python check_conformance.py V4/Public_ISO
```

- **Input:** Folder of `.xml` files (same as extractor).  
- **Output:** `conformance_report_<foldername>.xlsx` with:  
  - **Compliance Summary** — per file: conformant Yes/No, missing mandatory list, counts of present mandatory/conditional/optional.  
  - **Conformance Detail** — one row per file, one column per checked element; each cell is **Present**, **Empty**, or **Absent**; row 2 shows obligation (mandatory/conditional/optional).  
  - **Errors** (if any) — files skipped because they are not namespaced ISO 19139 (e.g. ArcGIS-style).

---

## Output Workbook

### Sheet: Compliance Summary (first sheet)

| Column | Description |
| ------ | ----------- |
| Filename | Source XML file name |
| ISO 19139 compliant | `Yes` if no mandatory fields are missing; otherwise `No` |
| Missing mandatory fields | Comma-separated list of mandatory attributes that are empty or absent |
| Missing count | Number of missing mandatory fields |

### Sheet: Metadata Export

- **Row 1:** Headers — `Filename` then all attribute names (alphabetically).  
- **Row 2:** Optionality — empty under `Filename`; under each attribute, one of `mandatory`, `optional`, `conditional`.  
- **Row 3+:** One row per XML file: filename in column A, then one cell per attribute (value or blank).  
- Header and optionality rows and the filename column are frozen for easier scrolling.

---

## ISO 19139 Optionality

The optionality row reflects a fixed mapping in the script (aligned with INSPIRE Regulation 1205/2008 and common ISO 19139 usage):

- **Mandatory:** e.g. Resource Title, Abstract, Topic Category, Keywords, geographic bounding box, Data Language, Scale Denominator, contact organisation/email/role, Distribution Online Resource Linkage, Lineage Statement, Data Quality scope, Metadata language/date/scope, Access/Other constraints, Conformance specification/pass, Use Limitation.  
- **Conditional:** e.g. Publication Date, Reference System Code/Code Space.  
- **Optional:** All other exported attributes (e.g. ArcGIS-specific, extra contact details).  

The mapping is defined in `FIELD_OBLIGATION` in `extract_metadata.py`; you can edit it to match your own profile or validator.

---

## Extracted Metadata (summary)

The script extracts (where present) attributes from:

- **Esri/ArcGIS:** Format, profile, dates, item name, content type, native extent, thumbnail, coordinate system, scale range.  
- **Data identification:** Title, alternate/collection title, abstract, extent description, geographic bounds, keywords, purpose, credit, use limitation, access/other constraints, language, character set, spatial representation type, scale denominator, environment, status, graphic overview, maintenance frequency, topic category, other keywords (with thesaurus name).  
- **Contact:** Individual/organisation/position name, address (delivery point, city, admin area, postal code, country), email, phone, online resource, hours, instructions, role.  
- **Spatial representation:** Topology level, geometry object type/count.  
- **Reference system:** Code, code space, version.  
- **Data quality:** Scope level, lineage statement, quality report type, conformance specification title, conformance pass.  
- **Distribution:** Online resource linkage, protocol, name, description.  
- **Metadata:** Maintenance frequency, language, country code, scope code, hierarchy level name, standard name/version, file ID, character set, date stamp.  
- **Spatial domain (Esri):** Feature name, type, geometry code.  
- **Entity/attributes (eainfo):** Entity type label/type/count, attribute names list.  

When the same logical field appears more than once (e.g. multiple keywords), values are concatenated with ` | `.

---

## Project Layout

```text
LandIS-ISO19139-Metadata-Compliance/
├── extract_metadata.py              # Main script: batch extract + Excel + compliance
├── check_conformance.py             # Strict namespace-aware ISO 19139 conformance checker
├── README.md                         # This file
├── LICENSE
├── .gitignore
├── metadata_export_<folder>.xlsx    # Generated by extract_metadata.py (ignored by git)
├── conformance_report_<folder>.xlsx # Generated by check_conformance.py (ignored by git)
├── venv/                             # Virtual environment (create via Installation; ignored by git)
├── xml/                              # Default input folder for XML files (create as needed)
│   └── *.xml
└── XML_Exports/                      # Example alternative input folder (create as needed)
    └── Public/
        └── *.xml
```

---

## Error Handling

- **Missing folder:** Script exits with an error if the configured XML folder does not exist.  
- **No XML files:** Reports “No XML files found” and exits without writing a workbook.  
- **Parse/processing error for one file:** The file is skipped, the error is printed, and the rest are processed. The workbook is still written using all successfully processed files.

---

## Customisation

- **Input folder:** Pass the folder as the first argument: `python extract_metadata.py <folder>`. Default is `xml`.  
- **Output path:** The output file is always `metadata_export_<foldername>.xlsx` in the current directory (foldername is the last component of the path). To change this, edit `parse_args()` in `extract_metadata.py`.  
- **Mandatory/optional/conditional:** Edit the `FIELD_OBLIGATION` dictionary in `extract_metadata.py`; the optionality row and Compliance Summary both use this mapping.

---

## License

This project is licensed under the **Creative Commons Attribution-ShareAlike 4.0 International** licence (CC BY-SA 4.0). You may share and adapt the work with appropriate credit and must distribute derivatives under the same licence. See [LICENSE](LICENSE) for details.
