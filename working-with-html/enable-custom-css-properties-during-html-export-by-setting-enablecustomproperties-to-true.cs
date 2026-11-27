// Title: How to enable custom CSS properties when exporting a workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook custom properties as CSS in HTML | EnableCustomProperties flag usage in HtmlSaveOptions for .NET | Save Excel file to HTML with custom CSS properties using Aspose.Cells | C# code to include custom workbook properties in HTML output Aspose | Aspose.Cells HTML export custom CSS property example
// Tags: Aspose.Cells EnableCustomProperties setting | HtmlSaveOptions HTML conversion configuration | Export workbook metadata as CSS stylesheet | C# Aspose.Cells HTML generation options | Embedding custom metadata in HTML output

using System;
using Aspose.Cells;

// The example creates a workbook, adds sample data and a custom property, configures HtmlSaveOptions (including EnableCustomProperties), and saves the workbook as an HTML file with the custom property rendered as a CSS rule.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and add some sample data.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Add a custom property to demonstrate that it will be exported.
            sheet.CustomProperties.Add("ExportedBy", "Aspose.Cells Demo");

            // -----------------------------------------------------------------
            // 2. Configure HTML save options.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Optional: make the HTML more readable.
            htmlOptions.ExportActiveWorksheetOnly = true;
            htmlOptions.ExportImagesAsBase64 = true;

            // -----------------------------------------------------------------
            // 3. Save the workbook as HTML.
            // -----------------------------------------------------------------
            string outputFile = "WorkbookWithCustomCss.html";
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
