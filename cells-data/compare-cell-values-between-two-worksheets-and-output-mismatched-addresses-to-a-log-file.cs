using System;
using System.IO;
using Aspose.Cells;

namespace CompareWorksheets
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the input workbook and the log file
            string workbookPath = "InputWorkbook.xlsx";
            string logPath = "MismatchedCells.log";

            // Load the workbook (create rule is the Workbook constructor, load rule is the Workbook(string) overload)
            Workbook workbook = new Workbook(workbookPath);

            // Access the first two worksheets to compare
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets[1];

            // Determine the maximum used rows and columns across both sheets
            int maxRow = Math.Max(sheet1.Cells.MaxDataRow, sheet2.Cells.MaxDataRow);
            int maxColumn = Math.Max(sheet1.Cells.MaxDataColumn, sheet2.Cells.MaxDataColumn);

            // Open a StreamWriter for the log file (output rule)
            using (StreamWriter writer = new StreamWriter(logPath))
            {
                // Iterate through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Retrieve cells; they may be null if never accessed
                        Cell cell1 = sheet1.Cells[row, col];
                        Cell cell2 = sheet2.Cells[row, col];

                        // Get the underlying values (could be null)
                        object value1 = cell1?.Value;
                        object value2 = cell2?.Value;

                        // Compare values, handling nulls correctly
                        bool areEqual = (value1 == null && value2 == null) ||
                                        (value1 != null && value1.Equals(value2));

                        // If values differ, write the address to the log
                        if (!areEqual)
                        {
                            // Use the address from the first sheet (e.g., "A1")
                            string address = cell1?.Name ?? CellsHelper.CellIndexToName(row, col);
                            writer.WriteLine(address);
                        }
                    }
                }
            }

            // Save the workbook (save rule)
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}