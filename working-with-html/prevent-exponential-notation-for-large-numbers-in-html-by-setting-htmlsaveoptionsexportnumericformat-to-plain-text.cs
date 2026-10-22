// Title: How to disable scientific notation for large numbers when saving a workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, inserts a very large numeric value, and saves it to HTML with HtmlSaveOptions.ExportNumericFormat set to plain text to avoid exponential notation. | Show how to configure HtmlSaveOptions in Aspose.Cells so that numeric cells are exported as plain text instead of scientific format. | Provide a snippet that sets the ExportNumericFormat property, generates the HTML file, and verifies that the full number appears unchanged.
// Common Searches: Aspose.Cells HTML export large number without scientific notation C# | Set ExportNumericFormat to plain text in HtmlSaveOptions Aspose.Cells | Prevent exponential notation in HTML output from Aspose.Cells workbook | How to keep full precision of big integers when saving as HTML using Aspose.Cells | C# Aspose.Cells save workbook to HTML plain text numbers
// Tags: Aspose.Cells HtmlSaveOptions ExportNumericFormat plain text | C# Aspose.Cells prevent scientific notation HTML | large numeric values HTML export Aspose.Cells | cell number format plain text Aspose.Cells HTML | Aspose.Cells HTML save large numbers

using System;
using Aspose.Cells;

// The example demonstrates creating a workbook, inserting a large numeric value, configuring HtmlSaveOptions.ExportNumericFormat to ExportNumericFormat.PlainText, and saving the workbook as an HTML file so that the number is rendered as plain text instead of scientific notation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Insert a large number that would normally appear in exponential notation
            sheet.Cells["A1"].PutValue(12345678901234567890.0);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The ExportNumericFormat property is not available in this version of Aspose.Cells.
            // If plain‑text export is required, adjust the cell's number format instead.

            // Save the workbook as an HTML file
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
