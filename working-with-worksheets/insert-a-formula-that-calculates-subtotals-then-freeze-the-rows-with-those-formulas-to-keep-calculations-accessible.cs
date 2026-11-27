// Title: Create grouped subtotal rows and freeze top rows in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: how to add grouped subtotal rows in an Excel file using Aspose.Cells C# | freeze top rows after inserting subtotals with Aspose.Cells .NET | Aspose.Cells Subtotal method example for region based sales data | C# code to freeze panes on the first two rows in a generated workbook | generate Excel report with subtotals and frozen header using Aspose.Cells
// Tags: Aspose.Cells subtotal rows C# | Aspose.Cells freeze panes .NET | grouped subtotal using Subtotal method | Excel worksheet header freeze Aspose.Cells | C# generate sales report with subtotals

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalFreezeDemo
{
    // The sample creates a new workbook, fills it with region and sales data, inserts subtotal rows for each region using the Subtotal method, freezes the top two rows (header and first subtotal) with FreezePanes, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "SubtotalAndFreezeDemo.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SalesData";

                // ------------------------------------------------------------
                // 2. Populate sample data
                // ------------------------------------------------------------
                // Header
                sheet.Cells["A1"].PutValue("Region");
                sheet.Cells["B1"].PutValue("Sales");

                // Sample rows (Region groups)
                string[] regions = { "North", "North", "South", "South", "East", "East", "West", "West" };
                double[] sales =   { 1200, 1500, 800, 950, 1100, 1300, 700, 850 };

                for (int i = 0; i < regions.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2 (1‑based Excel index)
                    sheet.Cells[row, 0].PutValue(regions[i]);   // Column A
                    sheet.Cells[row, 1].PutValue(sales[i]);     // Column B
                }

                // ------------------------------------------------------------
                // 3. Insert subtotal rows using the Subtotal method
                // ------------------------------------------------------------
                // Define the area that contains the data (including header)
                CellArea dataArea = CellArea.CreateCellArea("A1", $"B{regions.Length + 1}");

                // Parameters:
                //   dataArea          – the range that contains the data
                //   groupByColumnIndex – column index (0‑based) to group by (Region)
                //   function           – ConsolidationFunction.Sum for the subtotal
                //   subtotalColumnList – which columns to subtotal (Sales column = index 1)
                sheet.Cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 });

                // ------------------------------------------------------------
                // 4. Freeze the rows that contain the header and the first subtotal
                // ------------------------------------------------------------
                // After calling Subtotal, Aspose.Cells inserts a subtotal row after each group.
                // The first subtotal will appear after the first group (rows 2‑3). 
                // To keep the header and that first subtotal always visible, freeze the top 2 rows.
                sheet.FreezePanes(2, 0, 2, 0);

                // ------------------------------------------------------------
                // 5. Save the workbook
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The file contains subtotal rows and the top two rows are frozen.");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing
                Console.Error.WriteLine("An error occurred while creating the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
