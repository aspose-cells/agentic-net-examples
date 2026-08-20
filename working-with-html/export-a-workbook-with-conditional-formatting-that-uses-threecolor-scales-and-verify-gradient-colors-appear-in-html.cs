// Title: Export Workbook with 3‑Color Scale Conditional Formatting to HTML and Validate Gradient Colors (C# Aspose.Cells)
// Description: Creates a 10×10 multiplication table, applies a red‑yellow‑green three‑color scale conditional formatting to A1:J10, saves the workbook as HTML with separate CSS and retained unused styles, then reads the HTML file to confirm the presence of the gradient hex codes #FF0000, #FFFF00 and #008000.
// Keywords: Aspose.Cells | C# | HTML export | conditional formatting | three color scale | color scale gradient | HtmlSaveOptions | ExportWorksheetCSSSeparately | ExcludeUnusedStyles | verify HTML colors | gradient hex codes | web reporting
// Common Searches: Aspose.Cells export three color scale to HTML | C# verify gradient colors in exported HTML | keep color scale definitions when saving workbook as HTML Aspose.Cells | HtmlSaveOptions.ExcludeUnusedStyles effect on conditional formatting | read generated HTML and check hex colors Aspose.Cells
// Developer Intent: Generate an HTML file from a workbook that preserves a three‑color scale conditional formatting and programmatically confirm that the gradient colors are embedded in the output.
// Use Cases: Create web‑ready reports that display heat‑map style visualisation using a red‑yellow‑green scale. | Automated testing to ensure conditional formatting survives HTML conversion. | Produce dashboards with CSS‑separated output for easier styling and maintenance. | Validate migration of Excel workbooks to HTML for compliance or archival purposes.
// AI Prompts: Write C# code with Aspose.Cells to apply a red‑yellow‑green three‑color scale to range A1:J10 and save the workbook as HTML with separate CSS and unused styles retained. | Provide a C# method that reads the saved HTML file and returns true only if the hex codes #FF0000, #FFFF00, and #008000 are present. | Explain how HtmlSaveOptions.ExportWorksheetCSSSeparately and ExcludeUnusedStyles influence the export of gradient definitions for color scales. | Generate a unit‑test that asserts the presence of the three gradient colors in the HTML output produced by Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsColorScaleHtmlDemo
{
    // Creates a 10×10 multiplication table, applies a red‑yellow‑green three‑color scale conditional formatting to A1:J10, saves the workbook as HTML with separate CSS and retained unused styles, then reads the HTML file to confirm the presence of the gradient hex codes #FF0000, #FFFF00 and #008000.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (10x10 multiplication table)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue(row * col);
                }
            }

            // Add a three‑color scale conditional formatting to the range A1:J10
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 9
            };
            fcs.AddArea(area);

            // Add the ColorScale condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition fc = fcs[conditionIdx];

            // Configure the three‑color scale (Red → Yellow → Green)
            fc.ColorScale.Is3ColorScale = true;

            fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
            fc.ColorScale.MinColor = Color.Red;

            fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
            fc.ColorScale.MidCfvo.Value = 50;
            fc.ColorScale.MidColor = Color.Yellow;

            fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
            fc.ColorScale.MaxColor = Color.Green;

            // Save the workbook as HTML
            string htmlFile = "ColorScale.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export CSS separately to make it easier to inspect the generated styles
                ExportWorksheetCSSSeparately = true,
                // Keep unused styles to ensure all gradient definitions are present
                ExcludeUnusedStyles = false
            };
            workbook.Save(htmlFile, htmlOptions);

            // Verify that the gradient colors are present in the generated HTML
            string htmlContent = File.ReadAllText(htmlFile);

            bool hasRed = htmlContent.IndexOf("#FF0000", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasYellow = htmlContent.IndexOf("#FFFF00", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasGreen = htmlContent.IndexOf("#008000", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine("Verification of gradient colors in HTML:");
            Console.WriteLine($"Red color found:   {hasRed}");
            Console.WriteLine($"Yellow color found:{hasYellow}");
            Console.WriteLine($"Green color found: {hasGreen}");

            // Simple result output
            if (hasRed && hasYellow && hasGreen)
            {
                Console.WriteLine("All gradient colors are present in the HTML output.");
            }
            else
            {
                Console.WriteLine("One or more gradient colors are missing in the HTML output.");
            }
        }
    }
}
