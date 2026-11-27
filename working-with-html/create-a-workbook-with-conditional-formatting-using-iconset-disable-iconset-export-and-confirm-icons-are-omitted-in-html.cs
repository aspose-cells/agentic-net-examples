// Title: Generate an Excel workbook with traffic‑light IconSet conditional formatting and suppress icon export when saving to HTML using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide IconSet icons in HTML output | how to prevent conditional formatting icons from being exported to HTML with Aspose.Cells | save workbook as HTML without traffic light icons using HtmlSaveOptions ExportConditionalFormatting false | C# example disabling IconSet export in Aspose.Cells HTML conversion | verify that IconSet conditional formatting is omitted in generated HTML file Aspose.Cells
// Tags: IconSet conditional formatting HtmlSaveOptions | disable conditional formatting export Aspose.Cells | export Excel to HTML without icons C# | traffic lights IconSet Aspose.Cells | HTML output verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIconSetHtmlDemo
{
    // The sample creates a new workbook, applies a three‑traffic‑light IconSet conditional format to cells A1:A5, configures HtmlSaveOptions with ExportConditionalFormatting set to false to suppress icon markup, saves the file as HTML, reads the output, and confirms that no IconSet‑related HTML is present.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (values 10 to 50) in A1:A5
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 0].PutValue((i + 1) * 10);
                }

                // Add IconSet conditional formatting to range A1:A5 (3 traffic lights)
                int startRow = 0, endRow = 4, startColumn = 0, endColumn = 0;

                // Create a new conditional formatting collection
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the area to which the conditional formatting applies
                CellArea area = new CellArea
                {
                    StartRow = startRow,
                    StartColumn = startColumn,
                    EndRow = endRow,
                    EndColumn = endColumn
                };
                cfCollection.AddArea(area);

                // Add an IconSet condition
                int conditionIndex = cfCollection.AddCondition(ConditionType.IconSet);
                FormatCondition iconSetCondition = cfCollection[conditionIndex];
                iconSetCondition.IconSet = IconSetType.TrafficLights3;

                // Prepare HTML save options with conditional formatting disabled
                string htmlPath = "IconSetDemo.html";
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportConditionalFormatting = false // Disable IconSet export
                };

                // Save workbook as HTML
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML at: {Path.GetFullPath(htmlPath)}");

                // Verify that the HTML file exists before reading
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Verification FAILED: HTML file was not created.");
                    return;
                }

                // Read generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Simple heuristic: look for typical icon class names used by Aspose.Cells HTML export
                bool containsIcon = htmlContent.IndexOf("icon", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    htmlContent.IndexOf("cficon", StringComparison.OrdinalIgnoreCase) >= 0;

                if (containsIcon)
                {
                    Console.WriteLine("Verification FAILED: Icon markup was found in the HTML output.");
                }
                else
                {
                    Console.WriteLine("Verification PASSED: No icon markup found; IconSet export is disabled.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
