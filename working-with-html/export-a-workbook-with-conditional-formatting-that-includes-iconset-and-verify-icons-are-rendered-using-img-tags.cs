// Title: Export IconSet Conditional Formatting to HTML with <img> Icons using Aspose.Cells C#
// Description: Creates a workbook, fills column A with numbers, applies a TrafficLights31 IconSet to A1:A10, and saves the file as HTML with HtmlSaveOptions.ExportImagesAsBase64 set to false so each icon is written as an external image referenced by an <img> tag. The sample then reads the HTML to confirm the presence of the <img> elements.
// Keywords: Aspose.Cells | C# | HTML export | IconSet conditional formatting | TrafficLights31 | ExportImagesAsBase64 false | img tag verification | conditional formatting icons HTML | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells export IconSet to HTML | How to save conditional formatting icons as images in HTML | C# Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | Verify <img> tags in Aspose.Cells HTML output | IconSet HTML export example
// Developer Intent: Generate HTML from a workbook that includes IconSet conditional formatting and ensure the icons appear as separate image files referenced by <img> tags.
// Use Cases: Build a web dashboard where traffic‑light icons from an IconSet visually represent KPI values. | Automate a regression test that exports a spreadsheet to HTML and checks that each IconSet cell contains an <img> element pointing to the correct icon file. | Create email‑ready HTML reports with conditional‑formatting icons saved as external images for consistent rendering across mail clients.
// AI Prompts: Write C# code with Aspose.Cells to add a TrafficLights31 IconSet to a range and export the workbook to HTML using external image files for the icons. | Provide a method that parses the generated HTML and asserts that every cell containing the IconSet includes an <img> tag referencing the correct icon file. | Explain the impact of setting HtmlSaveOptions.ExportImagesAsBase64 to false on the rendering of conditional‑formatting icons in exported HTML.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills column A with numbers, applies a TrafficLights31 IconSet to A1:A10, and saves the file as HTML with HtmlSaveOptions.ExportImagesAsBase64 set to false so each icon is written as an external image referenced by an <img> tag. The sample then reads the HTML to confirm the presence of the <img> elements.
class IconSetHtmlExportDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate column A with sample numeric values
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10);
        }

        // Add an IconSet conditional formatting to the range A1:A10
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the cell area for the conditional formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add the IconSet condition and configure its type
        int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[conditionIndex];
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true; // optional: display the cell value alongside the icon

        // Prepare HTML save options to export images as separate files (so <img> tags are used)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = false // ensures icons are saved as image files referenced by <img>
        };

        // Define output paths
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);
        string htmlPath = Path.Combine(outputDir, "IconSet.html");

        // Save the workbook as HTML
        workbook.Save(htmlPath, htmlOptions);

        // Verify that the generated HTML contains <img> tags for the icons
        string htmlContent = File.ReadAllText(htmlPath);
        bool containsImgTag = htmlContent.Contains("<img");
        Console.WriteLine("HTML contains <img> tags for icons: " + containsImgTag);
    }
}
