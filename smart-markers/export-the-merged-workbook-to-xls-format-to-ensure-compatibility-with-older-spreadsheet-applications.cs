using Aspose.Cells;
using System;

class ExportToXls
{
    static void Main()
    {
        // Load the merged workbook (replace with your actual source file)
        Workbook workbook = new Workbook("merged.xlsx");

        // Create save options for the Excel 97-2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Optional: configure desired options
        saveOptions.MatchColor = true;               // Match font colors to the 56‑color palette
        saveOptions.ValidateMergedAreas = true;      // Validate merged cells before saving
        saveOptions.ClearData = false;               // Keep workbook data after saving

        // Save the workbook as an XLS file using the specified options
        workbook.Save("merged_output.xls", saveOptions);
    }
}