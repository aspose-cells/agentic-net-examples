// Title: How to autofit columns and rows after loading an HTML file into an Aspose.Cells workbook and export with TableCssId in C#
// Common Searches: C# Aspose.Cells autofit columns and rows after loading HTML workbook | Save Aspose.Cells workbook to HTML with a specific TableCssId | Preserve HTML table layout when exporting from Aspose.Cells using TableCssId | Load HTML into Aspose.Cells workbook and automatically adjust column widths | Aspose.Cells HtmlSaveOptions TableCssId example for .NET
// Tags: autofit columns rows Aspose.Cells HTML | HtmlSaveOptions TableCssId export | load HTML into Aspose.Cells workbook C# | preserve HTML table layout Aspose.Cells | Aspose.Cells autofit after HTML load

using Aspose.Cells;
using System;

// The example loads an HTML file into an Aspose.Cells Workbook, automatically autofits all columns and rows on the first worksheet, sets HtmlSaveOptions.TableCssId to retain the original table layout, and saves the workbook back to HTML.
class Program
{
    static void Main()
    {
        // Load the HTML file into a workbook
        var loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Get the first worksheet (the HTML is loaded into the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Autofit all columns in the worksheet
        sheet.AutoFitColumns();

        // Autofit all rows in the worksheet
        sheet.AutoFitRows();

        // Prepare HTML save options and set TableCssId to keep the layout
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.TableCssId = "myTable";

        // Export the workbook back to HTML with the specified TableCssId
        workbook.Save("output.html", saveOptions);
    }
}
