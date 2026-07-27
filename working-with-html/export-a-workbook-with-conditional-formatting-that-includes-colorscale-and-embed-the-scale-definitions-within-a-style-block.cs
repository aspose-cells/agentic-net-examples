// Title: C# – Export Excel with 3‑Color Scale Conditional Formatting to a Single HTML File and Embed CSS using Aspose.Cells
// Description: Demonstrates how to create a workbook, apply a red‑yellow‑green three‑color scale conditional format, and save it as a single HTML document with a custom <style> block via Aspose.Cells HtmlSaveOptions.
// Keywords: Aspose.Cells C# export HTML | ColorScale conditional formatting | HtmlSaveOptions SaveAsSingleFile | embed CSS in Aspose.Cells HTML | Excel to HTML with color scale | custom CSS style block Aspose | single file HTML output
// Common Searches: export Excel with color scale to HTML Aspose.Cells | embed custom CSS when saving workbook as HTML .NET | Aspose.Cells HtmlSaveOptions SaveAsSingleFile example | how to add a style block to Aspose.Cells HTML output | C# create 3‑color scale conditional formatting
// Developer Intent: Generate a self‑contained HTML file from an Excel workbook that includes a three‑color scale conditional format and custom CSS definitions.
// Use Cases: Produce web‑ready reports where cell backgrounds visualize data ranges. | Distribute Excel content without external style sheets, simplifying deployment. | Combine Aspose.Cells generated conditional‑formatting CSS with brand‑specific styles in one HTML file.
// AI Prompts: Write C# code that adds a red‑yellow‑green ColorScale conditional format to a worksheet and saves the workbook as a single HTML file with an embedded <style> block using Aspose.Cells. | Show how to configure HtmlSaveOptions to embed custom CSS while preserving Aspose.Cells' automatically generated conditional‑formatting classes. | Explain how to retrieve and customize the CSS class names created for a ColorScale condition in the HTML output of Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsColorScaleExport
{
    // Demonstrates how to create a workbook, apply a red‑yellow‑green three‑color scale conditional format, and save it as a single HTML document with a custom <style> block via Aspose.Cells HtmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (10x10 matrix)
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    worksheet.Cells[i, j].PutValue(i * j);
                }
            }

            // Add a ColorScale conditional formatting rule
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting applies
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 9
            };
            fcs.AddArea(area);

            // Add the ColorScale condition
            int conditionIndex = fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition fc = fcs[conditionIndex];

            // Configure a 3‑color scale: Red (min) → Yellow (mid) → Green (max)
            ColorScale cs = fc.ColorScale;
            cs.Is3ColorScale = true;

            // Minimum value (type = Min) – red
            cs.MinCfvo.Type = FormatConditionValueType.Min;
            cs.MinColor = Color.Red;

            // Mid value (type = Percentile, 50th percentile) – yellow
            cs.MidCfvo.Type = FormatConditionValueType.Percentile;
            cs.MidCfvo.Value = 50;
            cs.MidColor = Color.Yellow;

            // Maximum value (type = Max) – green
            cs.MaxCfvo.Type = FormatConditionValueType.Max;
            cs.MaxColor = Color.Green;

            // Prepare HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Save as a single HTML file so that the CSS block is embedded
            saveOptions.SaveAsSingleFile = true;

            // Embed custom CSS that describes the color scale (for illustration)
            // The class name "customColorScale" is not directly used by Aspose.Cells,
            // but demonstrates how to add a <style> block.
            saveOptions.CssStyles = @"
                .customColorScale {
                    background: linear-gradient(to right, red, yellow, green);
                }
                /* Aspose.Cells generated color scale classes will also be included automatically */
            ";

            // Export the workbook to HTML with the embedded style block
            string outputPath = "ColorScaleExport.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML with ColorScale. File: {outputPath}");
        }
    }
}
