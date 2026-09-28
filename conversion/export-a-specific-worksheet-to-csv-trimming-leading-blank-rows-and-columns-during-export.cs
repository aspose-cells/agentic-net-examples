// Title: Export a selected worksheet to CSV with trimmed leading blank rows and columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, activates a specific worksheet, and saves it as a CSV while removing any leading empty rows and columns with Aspose.Cells. | Show how to configure TxtSaveOptions for CSV output, enable TrimLeadingBlankRowAndColumn, and export only the active sheet in a .NET application.
// Common Searches: Aspose.Cells export only the active sheet to CSV and ignore leading empty rows | C# trim blank rows and columns when saving a worksheet as CSV using Aspose.Cells | How to use TxtSaveOptions to export a single worksheet to CSV in .NET
// Tags: Aspose.Cells active sheet CSV output | TxtSaveOptions TrimLeadingBlankRowAndColumn usage | C# export single worksheet as CSV | CSV conversion trimming empty rows Aspose.Cells | Save specific worksheet to CSV .NET

using System;
using Aspose.Cells;

// Loads an Excel workbook, activates a chosen worksheet, configures TxtSaveOptions with TrimLeadingBlankRowAndColumn and ExportAllSheets=false, and saves the active sheet as a CSV file.
class ExportWorksheetToCsv
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Index of the worksheet you want to export (0‑based)
        int worksheetToExport = 1; // example: second worksheet

        // Set the selected worksheet as the active one
        workbook.Worksheets.ActiveSheetIndex = worksheetToExport;

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            // Trim leading blank rows and columns like Excel does
            TrimLeadingBlankRowAndColumn = true,
            // Export only the active worksheet (default is false, but set explicitly for clarity)
            ExportAllSheets = false
        };

        // Save the active worksheet to a CSV file
        workbook.Save("exported_sheet.csv", csvOptions);
    }
}
