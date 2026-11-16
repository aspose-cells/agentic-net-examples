// Title: Export a workbook to HTML with a defined print area and without worksheet properties using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# export worksheet to HTML with specific print area | How to hide worksheet properties when saving Excel as HTML using Aspose.Cells | Set print area before converting Excel to HTML with Aspose.Cells in .NET | Save Excel workbook as HTML without worksheet metadata using Aspose.Cells C#
// Tags: HtmlSaveOptions ExportWorksheetProperties false | define print area Aspose.Cells | export workbook to HTML Aspose.Cells | omit worksheet properties HTML output | C# Aspose.Cells HTML conversion

using Aspose.Cells;
using System;

// The example creates a new Workbook, renames the first worksheet, sets its print area to A1:D20, configures HtmlSaveOptions to exclude worksheet properties, and saves the workbook as an HTML file named output.html.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Define the print area (for example, cells A1 to D20)
        sheet.PageSetup.PrintArea = "A1:D20";

        // Set HTML save options to omit worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorksheetProperties = false; // Do not include worksheet properties in the HTML

        // Export the workbook to an HTML file
        workbook.Save("output.html", htmlOptions);
    }
}
