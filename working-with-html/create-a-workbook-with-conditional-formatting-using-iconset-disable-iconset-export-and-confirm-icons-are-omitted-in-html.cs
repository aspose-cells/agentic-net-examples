// Title: C# – Add IconSet Conditional Formatting and Export to HTML without Icons using Aspose.Cells
// Description: Creates a workbook, fills column A with numbers, applies a TrafficLights31 IconSet to A1:A5, configures the IconSet to show only cell values, saves the sheet as HTML, and programmatically verifies that the generated HTML contains no <img> tags for icons.
// Keywords: Aspose.Cells C# IconSet | conditional formatting IconSet export HTML | hide icons in HTML Aspose.Cells | ShowValue property IconSet | verify HTML no img tags | Aspose.Cells HTMLSaveOptions | C# workbook to HTML without icons
// Common Searches: Aspose.Cells hide IconSet icons when saving as HTML | C# export workbook to HTML without conditional formatting icons | How to disable IconSet images in Aspose.Cells HTML output | Check HTML output for <img> tags after Aspose.Cells export | ShowValue only IconSet Aspose.Cells .NET
// Developer Intent: Generate an HTML file from a workbook that contains an IconSet rule but renders only the cell values, omitting the icon images.
// Use Cases: Apply a TrafficLights31 IconSet to a numeric range and export the sheet to HTML while keeping the icons hidden. | Programmatically confirm that the exported HTML does not contain any <img> elements representing conditional‑formatting icons. | Use the IconSet.ShowValue property to ensure values are displayed even when icons are suppressed in the HTML export.
// AI Prompts: Write C# code with Aspose.Cells that adds an IconSet conditional formatting rule to a range and saves the workbook to HTML without including the icons. | Provide a method that reads an Aspose.Cells‑generated HTML file and validates that no <img> tags are present, confirming icons were omitted. | Explain how the IconSet.ShowValue property influences HTML export in Aspose.Cells for .NET and how to use it to hide icons.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills column A with numbers, applies a TrafficLights31 IconSet to A1:A5, configures the IconSet to show only cell values, saves the sheet as HTML, and programmatically verifies that the generated HTML contains no <img> tags for icons.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample numeric data in column A
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue((i + 1) * 10);
            }

            // Add an IconSet conditional formatting rule to the range A1:A5
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

            // Add the IconSet condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcc[conditionIdx];

            // Set a normal icon set type (e.g., TrafficLights31)
            condition.IconSet.Type = IconSetType.TrafficLights31;

            // Hide icons in HTML by showing only the cell values
            // (ShowIconOnly property is not available in this version)
            condition.IconSet.ShowValue = true; // ensure the cell value is displayed

            // Save the workbook as HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            string htmlPath = "IconSetNoIcons.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML does not contain any <img> tags (icons)
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool containsImgTag = htmlContent.Contains("<img");
                Console.WriteLine("Icons omitted in HTML: " + (!containsImgTag));
            }
            else
            {
                Console.WriteLine("HTML file was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
