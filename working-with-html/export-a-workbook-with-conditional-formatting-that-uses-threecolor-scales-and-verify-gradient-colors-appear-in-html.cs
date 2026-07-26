// Title: Export Workbook with Three‑Color Scale Conditional Formatting to HTML and Verify Gradient Colors (C#)
// Description: Demonstrates how to create a workbook, fill A1:A10 with values, apply a red‑yellow‑green three‑color‑scale conditional format, save the file as HTML using Aspose.Cells for .NET, and programmatically confirm that the gradient colors appear in the generated HTML.
// Keywords: Aspose.Cells | C# export to HTML | conditional formatting | three color scale | color scale HTML | HtmlSaveOptions | gradient verification | heat map HTML | CellArea | FormatConditionType.ColorScale
// Common Searches: Aspose.Cells export three color scale to HTML | How to save conditional formatting colors in HTML with Aspose.Cells | Verify color scale gradient in generated HTML Aspose | C# Aspose.Cells HTMLSaveOptions keep styles | Check red yellow green colors in HTML output
// Developer Intent: Generate an HTML file from a workbook that retains a three‑color‑scale conditional format and programmatically confirm the red, yellow, and green gradient colors are present.
// Use Cases: Create heat‑map reports that can be viewed in browsers with accurate color‑scale rendering. | Automated testing of Aspose.Cells HTML export to ensure conditional formatting is preserved. | Produce web‑ready documentation of spreadsheets where color scales convey data insights.
// AI Prompts: Provide C# code using Aspose.Cells to apply a red‑yellow‑green three‑color scale to A1:A10 and save the workbook as HTML with separate CSS. | Show how to read the saved HTML file and assert that #FF0000, #FFFF00, and #008000 (or equivalent rgb values) exist. | Explain HtmlSaveOptions settings needed to retain all conditional‑formatting styles for inspection.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingHtmlDemo
{
    // Demonstrates how to create a workbook, fill A1:A10 with values, apply a red‑yellow‑green three‑color‑scale conditional format, save the file as HTML using Aspose.Cells for .NET, and programmatically confirm that the gradient colors appear in the generated HTML.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data (0..9) in a 10x1 range
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // 3. Add a three‑color scale conditional formatting to the range A1:A10
            int cfIndex = sheet.ConditionalFormattings.Add();                     // create a new ConditionalFormatting collection
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex]; // get the collection

            // Define the area the formatting applies to (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a ColorScale condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the three‑color scale (Red → Yellow → Green)
            condition.ColorScale.Is3ColorScale = true;
            condition.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
            condition.ColorScale.MinColor = Color.Red;

            condition.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
            condition.ColorScale.MidCfvo.Value = 50; // 50th percentile (median)
            condition.ColorScale.MidColor = Color.Yellow;

            condition.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
            condition.ColorScale.MaxColor = Color.Green;

            // 4. Save the workbook as HTML with separate CSS (easier to inspect)
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "ColorScaleDemo.html");
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportWorksheetCSSSeparately = true,
                ExcludeUnusedStyles = false   // keep all styles for verification
            };
            workbook.Save(htmlPath, htmlOptions);

            // 5. Verify that the generated HTML contains the gradient colors
            //    The colors may appear as hex strings (#FF0000, #FFFF00, #008000) or as rgb(...)
            string htmlContent = File.ReadAllText(htmlPath);

            bool containsRed = htmlContent.IndexOf("#FF0000", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               htmlContent.IndexOf("rgb(255,0,0)", StringComparison.OrdinalIgnoreCase) >= 0;

            bool containsYellow = htmlContent.IndexOf("#FFFF00", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                  htmlContent.IndexOf("rgb(255,255,0)", StringComparison.OrdinalIgnoreCase) >= 0;

            bool containsGreen = htmlContent.IndexOf("#008000", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 htmlContent.IndexOf("rgb(0,128,0)", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine("HTML file saved to: " + htmlPath);
            Console.WriteLine("Verification of gradient colors in HTML:");
            Console.WriteLine($" - Red color present:   {(containsRed ? "YES" : "NO")}");
            Console.WriteLine($" - Yellow color present: {(containsYellow ? "YES" : "NO")}");
            Console.WriteLine($" - Green color present: {(containsGreen ? "YES" : "NO")}");

            // Optional: indicate overall result
            if (containsRed && containsYellow && containsGreen)
                Console.WriteLine("All three gradient colors were found in the HTML output.");
            else
                Console.WriteLine("One or more gradient colors are missing in the HTML output.");
        }
    }
}
