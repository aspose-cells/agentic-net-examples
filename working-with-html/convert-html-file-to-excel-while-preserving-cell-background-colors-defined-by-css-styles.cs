// Title: Convert an HTML file to an XLSX workbook in C# while preserving CSS‑defined cell background colors with Aspose.Cells
// Common Searches: c# convert html to excel preserving css background colors using Aspose.Cells | how to load an html file with styles into an Aspose.Cells workbook | Aspose.Cells LoadOptions Html format keep cell formatting when saving as xlsx | retain CSS cell background colors after exporting html to xlsx in .NET | example of converting styled html table to Excel with Aspose.Cells C#
// Tags: html to xlsx conversion Aspose.Cells | css style import Aspose.Cells | LoadOptions Html format C# | Workbook.Save preserving styles | cell background color mapping Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads an HTML document into an Aspose.Cells Workbook using LoadOptions set to LoadFormat.Html, which parses CSS styles—including cell background colors—and then saves the workbook as an XLSX file, ensuring the original styling is retained.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Load the HTML file into a Workbook.
        // LoadOptions with LoadFormat.Html ensures that CSS styles,
        // including cell background colors, are interpreted and applied.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Path for the resulting Excel file
        string excelFilePath = "output.xlsx";

        // Save the workbook as an Excel file (XLSX format).
        // The background colors defined in the HTML/CSS are preserved.
        workbook.Save(excelFilePath, SaveFormat.Xlsx);
    }
}
