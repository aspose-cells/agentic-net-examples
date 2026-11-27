// Title: Create Excel subtotals for columns M and N together with Aspose.Cells for .NET
// Common Searches: how to generate subtotals for two columns in an Excel workbook using Aspose.Cells C# | Aspose.Cells C# subtotal multiple columns based on a grouping column | C# code example for Worksheet.Cells.Subtotal with columns M and N | apply sum subtotals to columns M and N while grouping by Category in Aspose.Cells | using CellArea to define range for subtotal of several columns in Aspose.Cells .NET
// Tags: Aspose.Cells Worksheet.Cells.Subtotal multiple columns | C# subtotal columns M N Aspose.Cells | Excel subtotal by grouping column Aspose.Cells | CellArea range definition for subtotal Aspose.Cells | ConsolidationFunction.Sum multiple columns .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // The sample creates a new workbook, fills column A with categories and columns M and N with random numbers, defines a CellArea covering A1:N11, groups rows by the Category column, and uses Worksheet.Cells.Subtotal with the Sum function to add subtotals for both columns M and N simultaneously before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = Path.Combine(Environment.CurrentDirectory, "SubtotalDemo.xlsx");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new Workbook and add sample data.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header row (row 0)
                cells[0, 0].PutValue("Category");   // Column A – grouping column
                cells[0, 12].PutValue("Amount M"); // Column M
                cells[0, 13].PutValue("Amount N"); // Column N

                // Populate 10 rows of sample data
                Random rnd = new Random();
                for (int i = 1; i <= 10; i++)
                {
                    // Alternate categories for grouping
                    string category = (i % 2 == 0) ? "Group A" : "Group B";
                    cells[i, 0].PutValue(category);

                    // Random numbers for columns M and N
                    cells[i, 12].PutValue(rnd.Next(100, 500));
                    cells[i, 13].PutValue(rnd.Next(100, 500));
                }

                // ------------------------------------------------------------
                // 2. Add subtotals for columns M (index 12) and N (index 13)
                //    using the Worksheet.Cells.Subtotal method.
                // ------------------------------------------------------------
                // Define the data area: from A1 (0,0) to N11 (10,13)
                CellArea dataArea = CellArea.CreateCellArea(0, 0, 10, 13);

                // Group by the first column (Category) – column index 0
                int groupByColumnIndex = 0;

                // Columns to subtotal: M and N
                int[] subtotalColumns = new int[] { 12, 13 };

                // Apply subtotal (SUM) for the selected columns
                cells.Subtotal(dataArea, groupByColumnIndex, ConsolidationFunction.Sum, subtotalColumns);

                // ------------------------------------------------------------
                // 3. Save the workbook with subtotals to a file.
                // ------------------------------------------------------------
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating subtotals:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
