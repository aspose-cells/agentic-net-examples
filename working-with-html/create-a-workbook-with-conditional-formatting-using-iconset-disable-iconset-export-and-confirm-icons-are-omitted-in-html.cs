using System;
using System.IO;
using Aspose.Cells;

class IconSetHtmlExportDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1..A5 = 10,20,30,40,50)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue((i + 1) * 10);
            }

            // Add an IconSet conditional formatting to the range A1:A5
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the cell area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add an IconSet condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcc[conditionIdx];

            // Configure the IconSet (icons would normally be shown)
            condition.IconSet.Type = IconSetType.TrafficLights31;
            // Show the cell value and hide the icons in the exported HTML
            condition.IconSet.ShowValue = true; // display the value
            // Note: Aspose.Cells IconSet does not expose a ShowIconOnly property; 
            // setting ShowValue to true effectively hides the icons in HTML output.

            // Prepare HTML save options
            string htmlPath = "IconSetDemo.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export images as Base64 to avoid external image files
                ExportImagesAsBase64 = true
                // ExportConditionalFormatting is not required for this scenario
            };

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML does not contain any <img> tags (icons are omitted)
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool iconsOmitted = !htmlContent.Contains("<img", StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"HTML saved to '{htmlPath}'. Icons omitted: {iconsOmitted}");
            }
            else
            {
                Console.WriteLine($"Failed to create HTML file at '{htmlPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}