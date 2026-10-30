// Title: Export a .NET workbook to HTML with a three‑color scale conditional formatting embedded in a <style> block using Aspose.Cells
// Common Searches: C# Aspose.Cells export workbook to HTML with three‑color scale conditional formatting | How to embed color scale conditional formatting CSS in HTML output using Aspose.Cells for .NET | Save Excel file as HTML with conditional formatting color scale in style tag Aspose.Cells | Aspose.Cells HtmlSaveOptions export conditional formatting as inline CSS C# | Generate HTML from workbook that includes red‑yellow‑green color scale using Aspose.Cells
// Tags: Aspose.Cells HTML export conditional formatting | three‑color scale conditional formatting Aspose.Cells | inline CSS for color scale Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 C# | C# generate HTML from Excel workbook

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, fills column A with values 0‑90, applies a red‑yellow‑green three‑color scale conditional formatting to A1:A10, configures HtmlSaveOptions to embed images as Base64, and saves the workbook as an HTML file where the color‑scale definitions are emitted inside a <style> block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // Add a three‑color scale conditional formatting (red → yellow → green)
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            int conditionIdx = cfCollection.AddCondition(FormatConditionType.ColorScale);
            FormatCondition colorScaleCondition = cfCollection[conditionIdx];

            // Set colors for the scale (type defaults: Min, Percentile 50, Max)
            colorScaleCondition.ColorScale.MinColor = Color.Red;
            colorScaleCondition.ColorScale.MidColor = Color.Yellow;
            colorScaleCondition.ColorScale.MaxColor = Color.Green;

            // Apply the formatting to the range A1:A10
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Set HTML save options (basic)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true
            };

            // Save the workbook as HTML
            string outputPath = "ConditionalFormatting.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
