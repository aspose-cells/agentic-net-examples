// Title: Extract Formatting‑Only Worksheet Names to a Text File with Aspose.Cells for .NET
// Description: Loads an Excel workbook using Aspose.Cells, scans each worksheet for cells that contain no values, collects the names of sheets that hold only formatting, and writes those names to a plain‑text file. Includes optional workbook save to demonstrate the full lifecycle.
// Keywords: Aspose.Cells C# worksheet detection | formatting only sheets extraction | list empty worksheets Aspose | export sheet names to text file | identify sheets without data .NET | Excel workbook analysis Aspose.Cells | C# save worksheet names txt
// Common Searches: how to find worksheets with only formatting using Aspose.Cells | save list of empty Excel sheets to a text file C# | Aspose.Cells detect sheets with no data values | extract worksheet names that contain no cell values | C# code to export formatting‑only sheet names
// Developer Intent: Detect worksheets that contain no cell values (formatting‑only) and write their names to a text file.
// Use Cases: Generate a quick audit report of formatting‑only sheets before publishing a workbook. | Skip empty worksheets during batch processing to improve performance. | Log sheet names that require data entry for quality‑control in automated pipelines.
// AI Prompts: Write C# code with Aspose.Cells that lists all worksheets without any cell values and saves the names to a CSV file. | Suggest performance optimizations for scanning large workbooks for formatting‑only sheets using Aspose.Cells. | Modify the example to also capture worksheets that contain only formulas but no constant values.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells, scans each worksheet for cells that contain no values, collects the names of sheets that hold only formatting, and writes those names to a plain‑text file. Includes optional workbook save to demonstrate the full lifecycle.
class FormattingOnlySheetsExtractor
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath); // load

        // List to hold names of worksheets that contain only formatting (no data)
        List<string> formattingOnlySheetNames = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool hasData = false;

            // Enumerate all cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell has a non‑null, non‑empty value
                if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                {
                    hasData = true;
                    break; // No need to continue scanning this sheet
                }
            }

            // If no data was found, the sheet contains only formatting
            if (!hasData)
            {
                formattingOnlySheetNames.Add(sheet.Name);
            }
        }

        // Write the collected worksheet names to a text file for further analysis
        string outputPath = "FormattingOnlySheets.txt";
        File.WriteAllLines(outputPath, formattingOnlySheetNames);

        // Optionally, save the workbook back (unchanged) to demonstrate the save lifecycle
        workbook.Save("output.xlsx"); // save
    }
}
