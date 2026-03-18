using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Text with a missing font");

        // Apply a style that uses a font which may not be installed
        Style style = worksheet.Cells["A1"].GetStyle();
        style.Font.Name = "NonExistingFont";
        worksheet.Cells["A1"].SetStyle(style);

        // Configure HTML save options and specify the default font to use
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DefaultFontName = "Arial"; // Fallback font when the original font is unavailable

        // Save the workbook as HTML using the specified options
        workbook.Save("output.html", saveOptions);
    }
}