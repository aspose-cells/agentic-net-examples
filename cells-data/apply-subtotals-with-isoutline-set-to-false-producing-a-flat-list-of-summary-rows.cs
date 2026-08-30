// Title: Create flat subtotal rows (no outline) with Aspose.Cells Cells.Subtotal in C#
// AI Prompts: Write C# code that uses Aspose.Cells to add subtotal rows grouped by a column and inserts the summary rows directly below the data (isOutline = false). | Show how to call Cells.Subtotal with ConsolidationFunction.Sum to generate a flat list of totals in an Excel worksheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add subtotal rows without outline | Cells.Subtotal flat list of totals example in .NET | Generate Excel subtotal rows below data using Aspose.Cells | C# Aspose.Cells subtotal grouped by column with isOutline false | Create summary rows in Excel workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells Cells.Subtotal flat list | C# Excel subtotal rows without outline | Aspose.Cells generate summary rows programmatically | ConsolidationFunction.Sum subtotal grouping | Excel workbook subtotal using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating sample data, and using Cells.Subtotal with isOutline set to false to add flat subtotal rows for the Sales column grouped by Region, then saving the workbook.
    public class SubtotalFlatListDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (A1:C6)
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

                object[,] data = new object[,]
                {
                    {"North", "Widget", 5000},
                    {"North", "Gadget", 3000},
                    {"South", "Widget", 6000},
                    {"South", "Gadget", 4000},
                    {"West",  "Widget", 4500}
                };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    cells[i + 1, 0].PutValue(data[i, 0]); // Region
                    cells[i + 1, 1].PutValue(data[i, 1]); // Product
                    cells[i + 1, 2].PutValue(data[i, 2]); // Sales
                }

                // Define the range that contains the data (A1:C6)
                CellArea area = CellArea.CreateCellArea("A1", "C6");

                // Apply subtotals:
                // - Group by the first column (Region) -> index 0
                // - Use SUM function
                // - Subtotal the Sales column -> index 2
                // - Settings produce a flat list of summary rows (no outline)
                cells.Subtotal(
                    area,
                    0,                                 // groupBy column index
                    ConsolidationFunction.Sum,         // subtotal function
                    new int[] { 2 },                   // columns to subtotal
                    false,                             // replace existing subtotals
                    false,                             // add page breaks between groups
                    false);                            // add summary below data (false => flat list)

                // Save the workbook
                string outputPath = "SubtotalFlatListDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            SubtotalFlatListDemo.Run();
        }
    }
}
