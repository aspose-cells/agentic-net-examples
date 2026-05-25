using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (A1:C6)
                worksheet.Cells["A1"].PutValue("Region");
                worksheet.Cells["B1"].PutValue("Product");
                worksheet.Cells["C1"].PutValue("Sales");

                object[,] data = new object[,]
                {
                    { "North", "Widget", 5000 },
                    { "North", "Gadget", 3000 },
                    { "South", "Widget", 6000 },
                    { "South", "Gadget", 4000 },
                    { "West",  "Widget", 4500 }
                };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    worksheet.Cells[i + 1, 0].PutValue(data[i, 0]); // Region
                    worksheet.Cells[i + 1, 1].PutValue(data[i, 1]); // Product
                    worksheet.Cells[i + 1, 2].PutValue(data[i, 2]); // Sales
                }

                // Define the range that contains the data (A1:C6)
                CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

                // Apply subtotals:
                // - Group by the first column (Region) -> index 0
                // - Use SUM function
                // - Subtotal the Sales column -> index 2
                // - replace = false (do not replace existing subtotals)
                // - pageBreaks = false (no page breaks between groups)
                // - summaryBelowData = false (produce a flat list of summary rows, not an outline)
                worksheet.Cells.Subtotal(
                    area,
                    0,                         // groupBy column index
                    ConsolidationFunction.Sum, // subtotal function
                    new int[] { 2 },           // columns to subtotal
                    false,                     // replace existing subtotals
                    false,                     // no page breaks
                    false);                    // flat list (no outline)

                // Save the workbook
                string outputPath = "FlatSubtotalDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}