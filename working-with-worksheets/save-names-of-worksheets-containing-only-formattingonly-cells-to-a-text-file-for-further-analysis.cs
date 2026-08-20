// Title: C# – List Worksheets Containing Only Formatting and Save Names to a Text File (Aspose.Cells)
// Description: Loads an Excel workbook with Aspose.Cells, scans each worksheet for cell values, identifies sheets that hold only formatting, and writes their names to a plain‑text file for further analysis.
// Keywords: Aspose.Cells | C# | formatting only worksheets | empty sheet detection | export sheet names | Excel workbook analysis | list worksheets without data
// Common Searches: Aspose.Cells find worksheets with only formatting | C# get names of empty Excel sheets | save worksheet names to text file Aspose | detect sheets without data in .NET | list formatting‑only worksheets programmatically
// Developer Intent: Identify worksheets that contain only formatting (no cell values) and write their names to a text file.
// Use Cases: Generate a report of formatting‑only sheets before publishing a workbook. | Skip non‑data worksheets during bulk processing to improve performance. | Create a task list for data entry teams to populate empty sheets. | Audit workbooks for unused or placeholder sheets during quality checks.
// AI Prompts: Provide C# code using Aspose.Cells to list worksheets that have no data and save the names to a CSV file. | Show how to modify the sample to also include sheets that contain only formulas but no constant values. | Suggest performance optimizations for detecting formatting‑only worksheets in a workbook with thousands of sheets.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, scans each worksheet for cell values, identifies sheets that hold only formatting, and writes their names to a plain‑text file for further analysis.
class FormattingOnlySheetsExtractor
{
    static void Main()
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold names of worksheets that contain only formatting (no data)
        List<string> formattingOnlySheetNames = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool hasData = false;

            // Enumerate all cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // If a cell has a non‑null, non‑empty value, the sheet contains data
                if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                {
                    hasData = true;
                    break; // No need to check further cells in this sheet
                }
            }

            // If no data was found, the sheet is formatting‑only
            if (!hasData)
            {
                formattingOnlySheetNames.Add(sheet.Name);
            }
        }

        // Save the list of sheet names to a text file for further analysis
        File.WriteAllLines("FormattingOnlySheets.txt", formattingOnlySheetNames);
    }
}
