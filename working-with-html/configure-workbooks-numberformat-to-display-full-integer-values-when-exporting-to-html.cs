// Title: How to set a custom integer number format in Aspose.Cells for .NET so HTML export shows full integer values
// Common Searches: Aspose.Cells .NET how to display full integer values in HTML export | prevent scientific notation when saving workbook as HTML using Aspose.Cells | set custom number format '0' for cells before HTML conversion in C# | apply integer style to cells to keep full digits in Aspose.Cells HTML output | configure HtmlSaveOptions to retain cell number formatting in Aspose.Cells
// Tags: Aspose.Cells custom number format for HTML export | integer display format in Aspose.Cells .NET | apply style with custom format to cells C# | HtmlSaveOptions preserve cell formatting Aspose.Cells | full integer values in exported HTML Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, inserts large integer values, defines a style with the custom number format "0" to force full integer display, applies the style to the cells, and saves the workbook as HTML using HtmlSaveOptions, ensuring the integers appear without scientific notation.
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

            // Insert integer values to demonstrate full integer display
            sheet.Cells["A1"].PutValue(123456789);
            sheet.Cells["A2"].PutValue(987654321);

            // Create a style with a custom number format that shows full integers (no scientific notation)
            Style intStyle = workbook.CreateStyle();
            intStyle.Custom = "0"; // "0" forces display of the entire integer without decimals

            // Apply the style to the cells containing integers
            sheet.Cells["A1"].SetStyle(intStyle);
            sheet.Cells["A2"].SetStyle(intStyle);

            // Configure HTML save options (default behavior exports formatted values)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
