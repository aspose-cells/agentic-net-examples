using System;
using System.IO;
using Aspose.Cells;

class CompareWorksheets
{
    static void Main()
    {
        // Load the workbook containing the two worksheets to compare
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first and second worksheets (adjust indices or names as needed)
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets[1];

        // Determine the maximum used row and column across both sheets
        int maxRow = Math.Max(sheet1.Cells.MaxDataRow, sheet2.Cells.MaxDataRow);
        int maxCol = Math.Max(sheet1.Cells.MaxDataColumn, sheet2.Cells.MaxDataColumn);

        // Open a log file to record mismatched cell addresses
        using (StreamWriter logWriter = new StreamWriter("mismatches.log"))
        {
            // Iterate through each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve cells from both worksheets
                    Cell cell1 = sheet1.Cells[row, col];
                    Cell cell2 = sheet2.Cells[row, col];

                    // Compare the underlying values, handling possible nulls
                    bool areEqual;
                    if (cell1 == null && cell2 == null)
                    {
                        areEqual = true;
                    }
                    else if (cell1 == null || cell2 == null)
                    {
                        areEqual = false;
                    }
                    else
                    {
                        areEqual = object.Equals(cell1.Value, cell2.Value);
                    }

                    // If values differ, write the cell address (A1 style) to the log
                    if (!areEqual)
                    {
                        // Both cells share the same address, use either one
                        string address = cell1.Name; // e.g., "B3"
                        logWriter.WriteLine(address);
                    }
                }
            }
        }

        // Optionally save the workbook after processing
        workbook.Save("output.xlsx");
    }
}