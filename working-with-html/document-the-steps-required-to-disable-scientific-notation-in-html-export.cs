// Title: Disable scientific notation in HTML export using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that applies a custom number format to a cell to force plain decimal output and saves the workbook as HTML with Aspose.Cells. | Show how to configure HTML export options so that numeric values are rendered without exponent notation during HTML export.
// Common Searches: Aspose.Cells C# export to HTML without scientific notation for large numbers | prevent exponent display in HTML output using Aspose.Cells | show plain decimal values in HTML export with Aspose.Cells | disable scientific notation when saving workbook as HTML in .NET | Aspose.Cells HTML export keep numeric formatting plain
// Tags: cell style plain integer Aspose.Cells | HTML export options plain decimal | suppress exponent display Aspose.Cells | C# export workbook to HTML without exponent | Aspose.Cells numeric formatting for HTML

using Aspose.Cells;
using System;

// The example creates a workbook, inserts a large numeric value, applies a cell style with the custom format "0" to force plain integer display, and saves the workbook as HTML using default HtmlSaveOptions, producing HTML output where the number appears without scientific notation.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Insert a numeric value that would normally be displayed in scientific notation
            sheet.Cells["A1"].PutValue(123456789012345.0);

            // 4. Apply a custom number format that forces plain decimal representation (no scientific notation)
            Style plainStyle = workbook.CreateStyle();
            plainStyle.Custom = "0"; // display as integer without exponent
            sheet.Cells["A1"].SetStyle(plainStyle);

            // 5. Configure HTML export options (default behavior preserves the formatted string)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // 6. Save the workbook as an HTML file with the specified options
            workbook.Save("ExportedWithoutScientificNotation.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
