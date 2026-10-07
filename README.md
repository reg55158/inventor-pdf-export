# Inventor PDF Export

An Autodesk Inventor iLogic rule that exports the drawings for an assembly to PDF and merges them into one file.

## What it does

1. Looks at the active assembly and every part and sub-assembly it references.
2. For each file named like `###-####...`, looks for a drawing (`.idw`) with the same name in the same folder.
3. Exports each drawing (all sheets) to `PDF Export\<drawing name>.pdf`, next to the assembly. Assembly drawings come first by default; set `assembliesFirst = False` to sort everything by number instead.
4. Merges the PDFs, using Adobe Acrobat, into `<assembly name>_Drawings.pdf` in the assembly folder.
5. Lists any parts or assemblies with no drawing, and any PDFs skipped because they were locked. The list is also saved as `PDF Export\_Missing drawings.txt`.

If an old PDF can't be overwritten (for example, it's open in a PDF viewer), you get a Retry / Skip / Cancel prompt.

## Requirements

- Autodesk Inventor with iLogic
- Adobe Acrobat (the full version, not Reader) for the merge step. To skip merging, delete step 3 in the code.

## Usage

1. Open the top-level assembly in Inventor.
2. Add `ExportDrawingsToPDF.vb` as an iLogic rule (external or in the document).
3. Run the rule.
