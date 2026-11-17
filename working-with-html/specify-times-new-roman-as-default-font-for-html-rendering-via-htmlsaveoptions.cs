// Title: Set Times New Roman as the default font for HTML export using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to change default font when saving workbook as HTML | set default font to Times New Roman in HtmlSaveOptions Aspose.Cells | C# export Excel to HTML with a specific default font using Aspose.Cells | HtmlSaveOptions DefaultFontName property example in .NET | Aspose.Cells render HTML with Times New Roman as fallback font
// Tags: HtmlSaveOptions default font configuration | Aspose.Cells HTML export custom font | C# set default font for HTML output | Aspose.Cells rendering options Times New Roman | export workbook to HTML with specific font

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, writes "Hello, World!" to cell A1, configures HtmlSaveOptions by setting DefaultFontName to "Times New Roman", and saves the workbook as an HTML file (output.html) while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, World!");

            // Configure HTML save options with the desired default font
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.DefaultFontName = "Times New Roman";

            // Save the workbook as an HTML file using the specified options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
