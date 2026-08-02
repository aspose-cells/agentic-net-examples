// Title: Export Aspose.Cells Workbook to HTML without Worksheet Properties (C#)
// Description: Shows how to save an Aspose.Cells workbook as HTML while disabling worksheet property export by setting HtmlSaveOptions.ExportWorksheetProperties to false.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportWorksheetProperties | C# | Excel to HTML conversion | omit worksheet properties | HTML export options | disable sheet metadata
// Common Searches: Aspose.Cells export HTML without worksheet properties C# | HtmlSaveOptions ExportWorksheetProperties false example | How to hide worksheet metadata in HTML export Aspose.Cells | C# save Excel as HTML without sheet settings | Remove worksheet properties from Aspose.Cells HTML output
// Developer Intent: Export a workbook to HTML while omitting worksheet properties.
// Use Cases: Create clean HTML reports for public websites without exposing Excel sheet settings. | Generate lightweight HTML versions of workbooks for documentation portals where metadata should be hidden. | Automate batch conversion of multiple workbooks to HTML to reduce file size and simplify the output.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to HTML using HtmlSaveOptions with ExportWorksheetProperties set to false and writes the result to a MemoryStream. | Explain which worksheet property elements are excluded from the HTML when ExportWorksheetProperties is false and the effect on the final file. | Combine HtmlSaveOptions settings such as ExportGridLines, ExportImages, and ExportWorksheetProperties to produce a customized HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Shows how to save an Aspose.Cells workbook as HTML while disabling worksheet property export by setting HtmlSaveOptions.ExportWorksheetProperties to false.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Configure HTML save options to omit worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportWorksheetProperties = false; // Disable exporting of worksheet properties

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output_without_worksheet_props.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML without worksheet properties.");
        }
    }
}
