// Title: Generate a data‑type summary worksheet that counts each CellValueType in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to iterate over the used range of a worksheet, tally the occurrences of each CellValueType, and write the results to a new sheet. | Enhance the program to calculate the total number of cells and add a column that shows the percentage share of each data type in the summary worksheet. | Adjust the logic to treat blank cells as CellValueType.IsEmpty and include them in the count report alongside other types.
// Common Searches: Aspose.Cells C# count how many numeric, string, date, boolean and error cells are in a worksheet | C# enumerate used cells and get CellValueType distribution with Aspose.Cells | Create a summary report of cell data types in an Excel file using Aspose.Cells for .NET
// Tags: enumerate cells CellValueType Aspose.Cells | count cell data types worksheet Aspose.Cells | create summary report Excel Aspose.Cells | calculate cell type percentages .NET | handle blank cells Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSummaryReport
{
    // The program builds a workbook, populates cells with numeric, string, DateTime, Boolean, blank, and error values, iterates through the used range to count each CellValueType, writes the counts (and optionally percentages) to a new sheet named "SummaryReport", and saves the file as DataTypeSummaryReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data of various types
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Numeric values
            cells["A1"].PutValue(123);
            cells["A2"].PutValue(45.67);

            // String values
            cells["B1"].PutValue("Hello");
            cells["B2"].PutValue("World");

            // DateTime value
            cells["C1"].PutValue(DateTime.Now);

            // Boolean value
            cells["C2"].PutValue(true);

            // Null (blank) cell - leave D1 empty

            // Error value (example: divide by zero)
            cells["D1"].Formula = "=1/0";

            // Calculate formulas so that error values are evaluated
            workbook.CalculateFormula();

            // Dictionary to hold counts of each CellValueType
            Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Enumerate cells within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // If the cell has never been instantiated, its type is Unknown; skip it
                    if (cell == null || cell.Type == CellValueType.IsUnknown)
                        continue;

                    // Increment count for the cell's type
                    if (typeCounts.ContainsKey(cell.Type))
                        typeCounts[cell.Type]++;
                    else
                        typeCounts[cell.Type] = 1;
                }
            }

            // Add a new worksheet to hold the summary report
            Worksheet reportSheet = workbook.Worksheets.Add("SummaryReport");
            Cells reportCells = reportSheet.Cells;

            // Write header
            reportCells["A1"].PutValue("Cell Value Type");
            reportCells["B1"].PutValue("Count");

            // Write the counts
            int reportRow = 1; // start from second row (index 1)
            foreach (var kvp in typeCounts)
            {
                reportCells[reportRow, 0].PutValue(kvp.Key.ToString());
                reportCells[reportRow, 1].PutValue(kvp.Value);
                reportRow++;
            }

            // Save the workbook
            workbook.Save("DataTypeSummaryReport.xlsx");
        }
    }
}
