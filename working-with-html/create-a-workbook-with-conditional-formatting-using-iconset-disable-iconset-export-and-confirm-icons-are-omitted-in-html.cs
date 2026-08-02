// Title: Disable IconSet Export in HTML with Aspose.Cells for .NET
// Description: Creates a workbook, applies an IconSet conditional format to A1:A5, disables icon rendering by setting IconSet.Type to None, saves the file as XLSX and HTML, and verifies that the generated HTML contains no <img> tags for icons.
// Keywords: Aspose.Cells | IconSet | conditional formatting | HTML export | disable icons | C# .NET | SaveFormat.Html | icon omission | Excel to HTML
// Common Searches: Aspose.Cells hide IconSet icons in HTML | disable IconSet export .NET | remove conditional formatting icons from HTML output | IconSet.Type None effect | verify HTML does not contain icon images Aspose
// Developer Intent: Add an IconSet rule, prevent its icons from being written to the HTML file, and programmatically confirm that the icons are absent.
// Use Cases: Produce clean HTML reports from Excel workbooks where visual icons are not desired. | Offer an Excel version with full conditional formatting while delivering a lightweight HTML preview without icons. | Automate compliance checks that ensure generated HTML does not embed unwanted image assets.
// AI Prompts: Generate C# code using Aspose.Cells to apply an IconSet to a range and suppress its icons when saving as HTML. | Write a method that scans an HTML string and returns true if any IconSet <img> elements are present. | Explain why setting IconSet.Type to None removes icons from the HTML export and how it differs from other IconSet types.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIconSetExportDemo
{
    // Creates a workbook, applies an IconSet conditional format to A1:A5, disables icon rendering by setting IconSet.Type to None, saves the file as XLSX and HTML, and verifies that the generated HTML contains no <img> tags for icons.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate some sample numeric data in column A (A1..A5 = 10,20,30,40,50)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue((i + 1) * 10);
                }

                // 3. Add an IconSet conditional formatting rule to the range A1:A5
                int cfIndex = sheet.ConditionalFormattings.Add();                     // create conditional formatting collection
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex]; // get the collection

                // Define the cell area for the rule (A1:A5)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };
                fcs.AddArea(area);

                // Add the IconSet condition
                int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
                FormatCondition condition = fcs[conditionIdx];

                // Set a visible icon set type first (e.g., TrafficLights31)
                condition.IconSet.Type = IconSetType.TrafficLights31;

                // ---- Disable IconSet export ----
                // By changing the type to None, the icons are not rendered.
                // This effectively omits icons when the workbook is saved to HTML.
                condition.IconSet.Type = IconSetType.None;

                // 4. Save the workbook as an Excel file (optional, just for reference)
                string excelPath = "IconSetDemo.xlsx";
                workbook.Save(excelPath, SaveFormat.Xlsx);

                // 5. Save the workbook as HTML with default options
                string htmlPath = "IconSetDemo.html";
                workbook.Save(htmlPath, SaveFormat.Html); // use SaveFormat.Html to avoid null option issues

                // 6. Verify that the generated HTML does NOT contain any icon images
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool containsIconImg = htmlContent.IndexOf("<img", StringComparison.OrdinalIgnoreCase) >= 0;

                    Console.WriteLine($"HTML file saved to: {Path.GetFullPath(htmlPath)}");
                    Console.WriteLine(containsIconImg
                        ? "Icons were found in the HTML output (unexpected)."
                        : "No icons detected in the HTML output – IconSet export successfully disabled.");
                }
                else
                {
                    Console.WriteLine($"HTML file was not created: {htmlPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
