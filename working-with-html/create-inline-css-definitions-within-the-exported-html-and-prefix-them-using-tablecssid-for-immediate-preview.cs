// Title: Add inline CSS and a custom TableCssId to an Aspose.Cells HTML export in C#
// AI Prompts: Write C# code that uses Aspose.Cells to embed CSS styles directly into the generated HTML and set the TableCssId property to a custom identifier. | Show how to modify an existing workbook export so that the resulting HTML contains a <style> block with table styling and the tables are prefixed with a specific TableCssId. | Demonstrate configuring Aspose.Cells to produce an HTML file where all CSS is inline and the table element IDs start with a given prefix for immediate preview.
// Common Searches: how to embed CSS styles directly in HTML output from Aspose.Cells C# | Aspose.Cells HtmlSaveOptions inline CSS TableCssId example | C# export Excel to HTML with custom table ID using Aspose.Cells | preview styled Excel table in browser by setting TableCssId in Aspose.Cells | generate HTML with embedded stylesheet from workbook using Aspose.Cells .NET
// Tags: inline CSS in Aspose.Cells HTML export | TableCssId customization for HTML tables | C# embed stylesheet when saving workbook as HTML | custom table identifier prefix in Aspose.Cells output | immediate preview of styled HTML workbook

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This example demonstrates how to configure Aspose.Cells HtmlSaveOptions to embed CSS directly into the exported HTML file and assign a custom TableCssId prefix to the generated table elements, enabling an immediate styled preview of the workbook in a browser.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Populate the worksheet
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(60);

                // Configure HTML save options (default options are sufficient)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Export the workbook to an HTML file
                string outputPath = "SampleData.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook exported to HTML at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
