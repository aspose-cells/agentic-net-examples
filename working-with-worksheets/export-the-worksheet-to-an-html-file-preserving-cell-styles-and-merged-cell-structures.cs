// Title: Export a specific worksheet to HTML with cell styles and merged cells using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export selected worksheet to HTML preserving formatting and merged cells | How to save only the active sheet as HTML with grid lines using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions keep cell styles when converting Excel to HTML in C# | Export Excel worksheet to HTML file while retaining merged cell structure in C# | Save workbook worksheet as HTML with styles using Aspose.Cells SaveFormat.Html
// Tags: Aspose.Cells HtmlSaveOptions export worksheet | preserve merged cells in HTML export | export active worksheet as HTML C# | retain cell formatting in HTML conversion | export grid lines Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetToHtmlApp
{
    // The example loads an existing Excel workbook (or creates one), selects a worksheet, configures HtmlSaveOptions to include grid lines and export only the active sheet, then saves the worksheet as an HTML file while preserving cell styles and merged cell structures.
    class ExportWorksheetToHtml
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                Workbook workbook;

                // Load workbook if it exists; otherwise create a simple workbook.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "SampleSheet";
                    ws.Cells["A1"].PutValue("Sample Text");
                    ws.Cells["B2"].PutValue(12345);
                }

                // Get the first worksheet (or any specific worksheet you need).
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export grid lines (optional).
                    ExportGridLines = true,

                    // Export only the active worksheet.
                    ExportActiveWorksheetOnly = true
                };

                // Set the active worksheet to the one we want to export.
                workbook.Worksheets.ActiveSheetIndex = worksheet.Index;

                // Save the selected worksheet to an HTML file.
                const string outputPath = "WorksheetExport.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
