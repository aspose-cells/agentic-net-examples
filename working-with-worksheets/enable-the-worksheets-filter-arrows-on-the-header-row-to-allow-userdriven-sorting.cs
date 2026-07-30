// Title: Enable AutoFilter arrows on a worksheet header using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds header and sample rows, sets the AutoFilter range to display filter arrows on the header, and saves the file as WorksheetWithFilterArrows.xlsx.
// Keywords: Aspose.Cells | C# | .NET | AutoFilter | filter arrows | Excel sorting | worksheet header | programmatic filter | Excel automation
// Common Searches: Aspose.Cells show filter arrows in Excel | C# enable AutoFilter for a range | add sorting dropdowns to Excel header with Aspose | programmatically apply AutoFilter in .NET | display Excel filter arrows using Aspose.Cells
// Developer Intent: Show filter arrows on the header row so end users can sort and filter data directly in Excel.
// Use Cases: Generate a sales report where users can filter by category or price. | Export database query results to an interactive Excel file with built‑in filtering. | Create a reusable template that automatically applies AutoFilter to dynamic data ranges.
// AI Prompts: Write C# code with Aspose.Cells to apply an AutoFilter to range A1:D20 and display filter arrows on the header. | Show how to set the AutoFilter.Range property in Aspose.Cells and save the workbook with visible filter arrows. | Provide an example of adding a custom AutoFilter criteria (e.g., values > 100) after enabling filter arrows using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds header and sample rows, sets the AutoFilter range to display filter arrows on the header, and saves the file as WorksheetWithFilterArrows.xlsx.
    public class EnableFilterArrowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header row
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Category");
                worksheet.Cells["C1"].PutValue("Price");

                // Populate some sample data
                worksheet.Cells["A2"].PutValue("Laptop");
                worksheet.Cells["B2"].PutValue("Electronics");
                worksheet.Cells["C2"].PutValue(1200);

                worksheet.Cells["A3"].PutValue("Shirt");
                worksheet.Cells["B3"].PutValue("Clothing");
                worksheet.Cells["C3"].PutValue(45);

                worksheet.Cells["A4"].PutValue("Phone");
                worksheet.Cells["B4"].PutValue("Electronics");
                worksheet.Cells["C4"].PutValue(800);

                worksheet.Cells["A5"].PutValue("Book");
                worksheet.Cells["B5"].PutValue("Stationery");
                worksheet.Cells["C5"].PutValue(20);

                // Apply an AutoFilter to the range that includes the header row.
                // This will display filter arrows on the header cells, allowing users to sort/filter.
                worksheet.AutoFilter.Range = "A1:C5";

                // Save the workbook to verify the filter arrows appear.
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WorksheetWithFilterArrows.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableFilterArrowsDemo.Run();
        }
    }
}
