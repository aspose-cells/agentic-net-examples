// Title: Create flat subtotal rows without outline using Aspose.Cells in C#
// Description: This example builds a workbook, adds Region/Product/Sales data, defines the range A1:C6, and calls Cells.Subtotal with isOutline = false and summaryBelowData = false. The worksheet outline is turned off so the subtotal rows appear as a flat list above each group, and the file is saved as FlatSubtotalDemo.xlsx.
// Keywords: Aspose.Cells subtotal C# | flat summary rows | disable outline Aspose.Cells | isOutline false | summary row above data | C# Excel subtotal example | regional sales total Aspose | Excel grouping without outline
// Common Searches: Aspose.Cells flat subtotal rows C# | how to turn off outline when adding subtotals in Aspose.Cells | subtotal method isOutline false .NET | place summary rows above detail Aspose.Cells | C# generate Excel subtotals without grouping symbols
// Developer Intent: Add subtotal rows as a flat list (summary above details) while keeping the worksheet outline disabled.
// Use Cases: Generate a sales report where each region’s total is shown directly above its transactions, eliminating expandable outline symbols. | Prepare an Excel sheet for PDF/CSV export where all rows stay visible but group totals remain clearly displayed. | Design a printable workbook with continuous rows and no outline markers, ensuring totals are easy to read on paper.
// AI Prompts: Write C# code with Aspose.Cells to add flat subtotal rows (summary above data) and disable the outline view. | Show how to extend the FlatSubtotalDemo to subtotal multiple numeric columns while keeping a flat list format. | Explain the impact of the Subtotal method’s isOutline parameter and the Outline.SummaryRowBelow property on row placement.

using System;
using System.IO;
using Aspose.Cells;

namespace FlatSubtotalDemo
{
    // This example builds a workbook, adds Region/Product/Sales data, defines the range A1:C6, and calls Cells.Subtotal with isOutline = false and summaryBelowData = false. The worksheet outline is turned off so the subtotal rows appear as a flat list above each group, and the file is saved as FlatSubtotalDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (Region, Product, Sales)
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

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
                    cells[i + 1, 0].PutValue(data[i, 0]); // Region
                    cells[i + 1, 1].PutValue(data[i, 1]); // Product
                    cells[i + 1, 2].PutValue(data[i, 2]); // Sales
                }

                // Define the range that contains the data (A1:C6)
                CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

                // Apply subtotals:
                // - Group by the first column (Region) -> index 0
                // - Use SUM function
                // - Subtotal the Sales column -> index 2
                // - Do NOT replace existing subtotals, do NOT add page breaks,
                //   and set summaryBelowData to false (flat list of summary rows)
                cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, false, false, false);

                // Ensure the outline does not display summary rows below detail
                worksheet.Outline.SummaryRowBelow = false;
                worksheet.Outline.SummaryColumnRight = false;

                // Determine output file path and ensure directory exists
                string outputFile = "FlatSubtotalDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
