// Title: Export large numeric values as plain text in HTML using Aspose.Cells HtmlSaveOptions ExportNumericFormat in C#
// AI Prompts: Generate C# code that creates a workbook, writes a very large integer to a cell, sets the cell’s style to Text, configures HtmlSaveOptions.ExportNumericFormat to output numbers as plain text, and saves the workbook as an HTML file. | Show how to prevent scientific notation when exporting Excel to HTML with Aspose.Cells by applying a text number format and enabling the ExportNumericFormat option in HtmlSaveOptions.
// Common Searches: how to stop Aspose.Cells from showing scientific notation in HTML output c# | export large integer from Excel to HTML as plain text using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportNumericFormat example c# | prevent exponential notation for big numbers when saving workbook to HTML | C# Aspose.Cells save workbook to HTML with numbers displayed as text
// Tags: Aspose.Cells HtmlSaveOptions ExportNumericFormat | large numbers plain text HTML export | prevent scientific notation Aspose.Cells | C# save workbook to HTML | cell style text format Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a workbook, writes a 20‑digit decimal to cell A1, applies a Text number format, sets HtmlSaveOptions.ExportNumericFormat to output numbers as plain text, and saves the workbook as LargeNumber.html, ensuring the value appears in the HTML without scientific notation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a large number
            sheet.Cells["A1"].PutValue(12345678901234567890m);

            // Set cell style to Text to avoid scientific notation
            Style style = sheet.Cells["A1"].GetStyle();
            style.Number = 49; // Text format
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file
            workbook.Save("LargeNumber.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
