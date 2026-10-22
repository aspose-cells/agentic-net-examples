// Title: Set chart legend entries to transparent (no fill) for all charts in an Excel template using Aspose.Cells for .NET
// AI Prompts: Loop over every worksheet and chart and make the legend text invisible by assigning a transparent color with Aspose.Cells. | Programmatically clear the fill of all chart legend entries before saving the workbook in C#. | Use Aspose.Cells to apply a no‑fill style to legend items of every chart in a template Excel file.
// Common Searches: how to make chart legend entries transparent in Aspose.Cells C# | remove legend fill from all charts in an Excel workbook programmatically | Aspose.Cells set legend entry font color to transparent for multiple charts | clear chart legend background in a template workbook using .NET
// Tags: Aspose.Cells legend entry no fill | bulk modify chart legends .NET | apply transparent legend style Aspose | process all worksheets charts Aspose.Cells | template Excel chart formatting C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads a template Excel workbook, iterates through each worksheet and its charts, accesses each chart's legend, and sets the legend entry font color to transparent to simulate no fill, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWorkbook.xlsx";
            const string outputPath = "CustomizedWorkbook.xlsx";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    Legend legend = chart.Legend;
                    if (legend != null)
                    {
                        // Iterate through each legend entry
                        foreach (LegendEntry entry in legend.LegendEntries)
                        {
                            // Simulate no fill by making the legend entry font transparent
                            entry.Font.Color = Color.Transparent;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
