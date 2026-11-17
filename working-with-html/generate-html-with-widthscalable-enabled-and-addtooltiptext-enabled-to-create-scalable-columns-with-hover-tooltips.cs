// Title: Create HTML from an Excel workbook with scalable column widths and hover tooltips using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML with column width scaling enabled | How to include cell comments as hover tooltips in HTML output with Aspose.Cells | Enable WidthScalable and AddTooltipText in HtmlSaveOptions for .NET | Generate responsive HTML table from workbook using Aspose.Cells HtmlSaveOptions | C# example for scalable columns and tooltips when saving Excel as HTML
// Tags: HtmlSaveOptions WidthScalable property | HtmlSaveOptions AddTooltipText option | export Excel workbook to HTML with scalable columns | cell comments as HTML hover tooltips Aspose.Cells | responsive HTML table generation from Excel .NET

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates creating a workbook, populating data, and saving it as an HTML file with column width scaling and cell comment tooltips enabled via HtmlSaveOptions in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.50);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Enable column width scaling so columns adjust proportionally in the HTML output
                    WidthScalable = true,

                    // Add tooltip text (cell comments) to HTML cells for hover information
                    AddTooltipText = true
                };

                // Save the workbook as an HTML file with the specified options
                string outputFile = "ScalableColumnsWithTooltips.html";
                workbook.Save(outputFile, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
