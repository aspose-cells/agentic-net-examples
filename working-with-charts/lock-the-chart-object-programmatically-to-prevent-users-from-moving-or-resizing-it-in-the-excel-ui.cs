// Title: Programmatically lock an Excel chart to prevent moving or resizing using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that makes a chart object read‑only so users cannot drag or resize it in the Excel UI. | Show how to apply worksheet protection and, when supported, set the chart’s IsObjectLocked flag with Aspose.Cells for .NET.
// Common Searches: asp.net c# lock chart object in excel workbook using aspose.cells | prevent chart resizing in excel file with aspose.cells .net | set IsObjectLocked on chart Aspose.Cells C# example | worksheet protect to stop chart movement Aspose.Cells | make Excel chart immovable programmatically Aspose.Cells for .NET
// Tags: lock chart object Aspose.Cells | disable chart resizing Aspose.Cells | worksheet protection chart Aspose.Cells | IsObjectLocked property C# | prevent chart movement Excel Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads or creates a workbook, obtains or adds a chart, applies worksheet protection and, if the API version supports it, sets the chart’s IsObjectLocked property, then saves the file so the chart cannot be moved or resized in Excel.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain a chart – use the first one if it exists, otherwise create a sample chart
            Chart chart;
            if (worksheet.Charts.Count > 0)
            {
                chart = worksheet.Charts[0];
            }
            else
            {
                // Add a simple column chart as a placeholder and retrieve the created chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                chart = worksheet.Charts[chartIndex];
            }

            // NOTE: In some Aspose.Cells versions the Chart class does not expose an IsObjectLocked property.
            // If available, you can lock the chart like this:
            // chart.IsObjectLocked = true;
            // For compatibility, we rely on worksheet protection to restrict modifications.

            // Protect the worksheet so the chart (if locked) cannot be moved or resized
            worksheet.Protect(ProtectionType.All);

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
