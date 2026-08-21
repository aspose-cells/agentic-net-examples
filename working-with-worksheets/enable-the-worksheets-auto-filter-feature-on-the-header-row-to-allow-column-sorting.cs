// Title: Enable AutoFilter on a Header Row with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add header titles and sample data, set the AutoFilter range to A1:C4, flag the range as having headers, and save the file as AutoFilterEnabled.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells AutoFilter C# | enable AutoFilter Aspose.Cells | AutoFilter header row .NET | Excel column sorting Aspose.Cells | Aspose.Cells worksheet filter example
// Common Searches: Aspose.Cells enable AutoFilter on range | C# set AutoFilter headers Aspose.Cells | How to add Excel AutoFilter with Aspose.Cells | AutoFilter sorter HasHeaders property example
// Developer Intent: Add an AutoFilter to the worksheet’s header row so end users can sort and filter column data.
// Use Cases: Create a sales report where the first row contains column titles and users can filter by product, category, or price. | Export database query results to Excel with AutoFilter pre‑enabled for quick analysis. | Generate a dynamic data sheet that automatically applies AutoFilter to the populated range for interactive end‑user exploration.
// AI Prompts: Show C# code to enable AutoFilter on a worksheet range and set HasHeaders to true using Aspose.Cells. | Provide an Aspose.Cells for .NET example that applies AutoFilter to a range that expands to the last used row. | Explain how to customize AutoFilter dropdown options after enabling it with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFilterDemo
{
    // Demonstrates how to create a workbook, add header titles and sample data, set the AutoFilter range to A1:C4, flag the range as having headers, and save the file as AutoFilterEnabled.xlsx using Aspose.Cells for .NET.
    public class EnableAutoFilter
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

                // Populate sample data
                worksheet.Cells["A2"].PutValue("Laptop");
                worksheet.Cells["B2"].PutValue("Electronics");
                worksheet.Cells["C2"].PutValue(1200);

                worksheet.Cells["A3"].PutValue("Shirt");
                worksheet.Cells["B3"].PutValue("Clothing");
                worksheet.Cells["C3"].PutValue(45);

                worksheet.Cells["A4"].PutValue("Phone");
                worksheet.Cells["B4"].PutValue("Electronics");
                worksheet.Cells["C4"].PutValue(800);

                // Enable AutoFilter on the range A1:C4
                worksheet.AutoFilter.Range = "A1:C4";

                // Indicate that the range has headers
                worksheet.AutoFilter.Sorter.HasHeaders = true;

                // Save the workbook
                workbook.Save("AutoFilterEnabled.xlsx");
                Console.WriteLine("Workbook saved as AutoFilterEnabled.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableAutoFilter.Run();
        }
    }
}
