using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIconSetHtmlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample numeric data in column A (A1:A6)
                for (int i = 0; i < 6; i++)
                {
                    sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...
                }

                // Add an IconSet conditional formatting to the range A1:A6
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the cell area for the conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 5,
                    StartColumn = 0,
                    EndColumn = 0
                };
                fcc.AddArea(area);

                // Add the IconSet condition and set its type (e.g., TrafficLights31)
                int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
                FormatCondition condition = fcc[conditionIndex];
                condition.IconSet.Type = IconSetType.TrafficLights31;
                condition.IconSet.ShowValue = true; // optional: display cell value alongside the icon

                // Prepare HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export images (icons) as Base64 so they appear inside <img> tags
                    ExportImagesAsBase64 = true,
                    // Keep CSS inline for simplicity
                    ExportWorksheetCSSSeparately = false
                    // Note: ExportConditionalFormatting is not required for icon sets; it is not a valid property in this version
                };

                // Define output paths
                string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
                Directory.CreateDirectory(outputDir);
                string htmlPath = Path.Combine(outputDir, "IconSetDemo.html");

                // Save the workbook as HTML
                workbook.Save(htmlPath, htmlOptions);

                // Verify that the generated HTML contains <img> tags for the icons
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool containsImgTag = htmlContent.Contains("<img");
                    bool containsBase64 = htmlContent.Contains("data:image/png;base64");

                    Console.WriteLine("HTML file generated at: " + htmlPath);
                    Console.WriteLine("Verification result:");
                    Console.WriteLine(" - Contains <img> tag: " + containsImgTag);
                    Console.WriteLine(" - Contains Base64 image data: " + containsBase64);
                }
                else
                {
                    Console.WriteLine("Failed to generate HTML file at: " + htmlPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}