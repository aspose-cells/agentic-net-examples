// Title: C# – Remove Blank Rows & Columns with Aspose.Cells and Save as ODS
// Description: Shows how to build a workbook, insert data with intentional gaps, delete empty rows and columns via Cells.DeleteBlankRows and Cells.DeleteBlankColumns, set OdsSaveOptions, and export the cleaned workbook to OpenDocument Spreadsheet (ODS) format.
// Keywords: Aspose.Cells | C# | DeleteBlankRows | DeleteBlankColumns | remove empty rows | remove empty columns | ODS export | OdsSaveOptions | clean workbook | OpenDocument Spreadsheet | Excel to ODS conversion | data cleanup
// Common Searches: Aspose.Cells delete blank rows C# | How to remove empty columns before ODS export | Cells.DeleteBlankRows and DeleteBlankColumns example | Save cleaned workbook as ODS using Aspose.Cells | C# remove blank rows and columns Aspose
// Developer Intent: Eliminate all blank rows and columns from a workbook and then export it to ODS.
// Use Cases: Strip placeholder rows/columns from a template before generating a final ODS report. | Compact data imported from external systems to avoid unnecessary empty space in ODS files. | Prepare spreadsheets for downstream processing where blank rows or columns cause parsing errors.
// AI Prompts: Generate C# code that removes blank rows and columns from every worksheet in an Aspose.Cells workbook and saves it as ODS with custom OdsSaveOptions. | Explain the impact of DeleteBlankRows/DeleteBlankColumns on merged cells and how to preserve formatting before ODS conversion. | Provide a step‑by‑step tutorial for cleaning a workbook of empty rows/columns and exporting it to ODS using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsRemoveBlanks
{
    // Shows how to build a workbook, insert data with intentional gaps, delete empty rows and columns via Cells.DeleteBlankRows and Cells.DeleteBlankColumns, set OdsSaveOptions, and export the cleaned workbook to OpenDocument Spreadsheet (ODS) format.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with data and intentional blank rows/columns
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row1");
            // Row 3 is blank
            cells["A4"].PutValue("Row2");
            // Column B is blank
            cells["C1"].PutValue("ColC Header");
            cells["C2"].PutValue("ColC Data");

            // Remove all blank rows (method from Cells)
            cells.DeleteBlankRows();

            // Remove all blank columns (method from Cells)
            cells.DeleteBlankColumns();

            // Prepare ODS save options (optional, can use defaults)
            OdsSaveOptions odsOptions = new OdsSaveOptions();

            // Save the cleaned workbook as ODS (lifecycle: save)
            workbook.Save("CleanedWorkbook.ods", odsOptions);
        }
    }
}
