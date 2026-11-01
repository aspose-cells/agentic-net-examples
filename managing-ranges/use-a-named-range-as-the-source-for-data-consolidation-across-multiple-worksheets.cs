// Title: Consolidate numeric data from multiple worksheets into a single sheet using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells sum values across several worksheets into one sheet | how to aggregate data from multiple Excel sheets with Aspose.Cells .NET | example of consolidating worksheet ranges into a new worksheet using Aspose.Cells | loop through cells to combine sheet data in Aspose.Cells C# | save a consolidated workbook after summing ranges with Aspose.Cells
// Tags: worksheet data aggregation Aspose.Cells | cross-sheet summation C# Aspose.Cells | consolidated worksheet creation Aspose.Cells | cell iteration for data consolidation .NET | export consolidated workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace ConsolidationExample
{
    // The example creates a workbook with three worksheets containing numeric values, adds a fourth worksheet, iterates over the A1:B2 range on each source sheet, sums corresponding cell values, writes the totals to the consolidation sheet, and saves the result as ConsolidatedResult.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add two more worksheets
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Populate sample data in each worksheet
                // Sheet1 data
                sheet1.Cells["A1"].PutValue(1);
                sheet1.Cells["A2"].PutValue(2);
                sheet1.Cells["B1"].PutValue(10);
                sheet1.Cells["B2"].PutValue(20);

                // Sheet2 data
                sheet2.Cells["A1"].PutValue(3);
                sheet2.Cells["A2"].PutValue(4);
                sheet2.Cells["B1"].PutValue(30);
                sheet2.Cells["B2"].PutValue(40);

                // Sheet3 data
                sheet3.Cells["A1"].PutValue(5);
                sheet3.Cells["A2"].PutValue(6);
                sheet3.Cells["B1"].PutValue(50);
                sheet3.Cells["B2"].PutValue(60);

                // Add a new worksheet that will hold the consolidation result
                Worksheet consolidatedSheet = workbook.Worksheets.Add("Consolidated");

                // Manually consolidate (sum) data from the three sheets
                // Assuming the range A1:B2 in each sheet
                for (int row = 0; row < 2; row++)          // rows 0 and 1 correspond to A and B rows
                {
                    for (int col = 0; col < 2; col++)      // columns 0 and 1 correspond to columns A and B
                    {
                        double sum = 0;

                        // Sum values from each source sheet
                        sum += GetNumericValue(sheet1.Cells[row, col]);
                        sum += GetNumericValue(sheet2.Cells[row, col]);
                        sum += GetNumericValue(sheet3.Cells[row, col]);

                        // Put the summed value into the consolidated sheet
                        consolidatedSheet.Cells[row, col].PutValue(sum);
                    }
                }

                // Save the workbook to a file
                string outputPath = "ConsolidatedResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to safely retrieve numeric values from cells
        private static double GetNumericValue(Cell cell)
        {
            if (cell == null) return 0;
            object val = cell.Value;
            if (val == null) return 0;

            // Try to convert to double; if conversion fails, treat as zero
            if (double.TryParse(val.ToString(), out double result))
                return result;
            return 0;
        }
    }
}
