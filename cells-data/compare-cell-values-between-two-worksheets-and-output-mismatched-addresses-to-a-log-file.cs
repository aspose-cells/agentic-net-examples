// Title: C# – Compare Two Worksheets with Aspose.Cells and Log Mismatched Cell Addresses
// Description: Loads an Excel file, compares each cell of the first two worksheets using Aspose.Cells, writes the A1‑style addresses of all mismatched cells to a log file, and saves the workbook.
// Keywords: Aspose.Cells compare worksheets | C# Excel cell comparison | log mismatched cells | Excel diff report | A1 notation address | null‑safe value comparison | save workbook Aspose
// Common Searches: Aspose.Cells compare two sheets C# | log cell differences Aspose.Cells | find mismatched cells in Excel using .NET | Excel worksheet diff script | write Excel diff to log file
// Developer Intent: Detect cells whose values differ between two worksheets and record their addresses for review.
// Use Cases: Validate data integrity between source and target sheets in an ETL pipeline. | Create a diff report for financial reconciliation by listing changed cells. | Automate quality checks for template compliance by flagging non‑matching values.
// AI Prompts: Generate a C# method with Aspose.Cells that returns a list of mismatched cell addresses between two worksheets. | Show how to extend the example to include the original and new values beside each address in the log. | Explain how to compare formula results rather than the formula text using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, compares each cell of the first two worksheets using Aspose.Cells, writes the A1‑style addresses of all mismatched cells to a log file, and saves the workbook.
class CompareWorksheets
{
    static void Main()
    {
        // Load a workbook that contains the two worksheets to compare
        Workbook workbook = new Workbook("input.xlsx"); // create/load

        // Access the first two worksheets (index 0 and 1)
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets[1];

        // Determine the maximum rows and columns that contain data in either sheet
        int maxRow = Math.Max(sheet1.Cells.MaxDataRow, sheet2.Cells.MaxDataRow);
        int maxCol = Math.Max(sheet1.Cells.MaxDataColumn, sheet2.Cells.MaxDataColumn);

        // Open a log file to record mismatched cell addresses
        using (StreamWriter logWriter = new StreamWriter("mismatches.log"))
        {
            // Iterate through each cell within the determined range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve cells from both worksheets
                    Cell cell1 = sheet1.Cells[row, col];
                    Cell cell2 = sheet2.Cells[row, col];

                    // Get the underlying values (could be null)
                    object value1 = cell1?.Value;
                    object value2 = cell2?.Value;

                    // Compare values, handling nulls safely
                    bool areEqual = (value1 == null && value2 == null) ||
                                    (value1 != null && value1.Equals(value2));

                    // If values differ, write the cell address (A1 notation) to the log
                    if (!areEqual)
                    {
                        string address = CellsHelper.CellIndexToName(row, col);
                        logWriter.WriteLine(address);
                    }
                }
            }
        }

        // Save the workbook (no modifications made, but required by lifecycle rule)
        workbook.Save("output.xlsx");
    }
}
