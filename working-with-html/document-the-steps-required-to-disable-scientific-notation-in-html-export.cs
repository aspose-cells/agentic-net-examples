// Title: Disable scientific notation when exporting Excel to HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook to HTML with Aspose.Cells while applying a custom number format to prevent scientific notation for large and small numeric values. | Explain how to configure HtmlSaveOptions and cell styles in Aspose.Cells so that numbers appear in plain decimal form in the generated HTML file.
// Common Searches: asp.net aspose.cells html export prevent scientific notation for numbers | c# keep large numbers from showing as 1.23e+14 in HTML output using Aspose.Cells | set number format in Aspose.Cells to avoid exponent notation when saving as HTML | disable exponent display in Aspose.Cells HTMLSaveOptions example | export Excel to HTML without scientific notation Aspose.Cells C# tutorial
// Tags: Aspose.Cells number format for HTML export | HTMLSaveOptions disable scientific notation Aspose.Cells | prevent exponent notation in HTML output C# | cell style plain decimal format Aspose.Cells | large number display without scientific notation Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes a large integer and a tiny decimal, applies a custom numeric style ("0.###############") to both cells, configures HtmlSaveOptions, and saves the workbook as HTML, ensuring the values are rendered without scientific notation.
class DisableScientificNotationHtmlExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with values that would normally be shown in scientific notation
            sheet.Cells["A1"].PutValue(123456789012345L);   // Large integer
            sheet.Cells["A2"].PutValue(0.00000012345);    // Small decimal

            // Apply a custom number format that avoids scientific notation
            Style style = workbook.CreateStyle();
            style.Custom = "0.###############";
            sheet.Cells["A1"].SetStyle(style);
            sheet.Cells["A2"].SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file with the specified options
            string outputPath = "ExportedWithoutScientificNotation.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
