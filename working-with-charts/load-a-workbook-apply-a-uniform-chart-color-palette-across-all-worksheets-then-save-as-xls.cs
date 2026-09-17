// Title: Apply a uniform custom color palette to all chart series in every worksheet of an Excel workbook using Aspose.Cells for .NET and save as XLSX
// AI Prompts: Load an existing .xlsx file with Aspose.Cells for .NET, define a custom Color[] palette, iterate through each worksheet and chart, assign each series' Area.ForegroundColor and Border.Color from the palette, then save the workbook as a new .xlsx file. | Write a C# method that takes input and output paths, opens the workbook, applies a predefined six‑color palette to every chart series across all worksheets using Aspose.Cells, and returns the location of the saved file. | Create a script that verifies an Excel file's existence, opens it with Aspose.Cells, updates all chart series to use matching fill and border colors from a custom palette, handles any exceptions, and writes the modified workbook back to disk.
// Common Searches: how to change chart series colors for all worksheets using Aspose.Cells in C# | apply same color scheme to multiple charts in an Excel file with Aspose.Cells .NET | set custom palette for chart series programmatically Aspose.Cells | C# code to iterate over charts and update series fill color in a workbook | save workbook after modifying chart colors with Aspose.Cells for .NET
// Tags: Aspose.Cells set chart series color | C# apply custom chart palette | iterate worksheets charts Aspose.Cells | save workbook as XLSX Aspose.Cells | uniform chart color scheme .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX workbook, defines a six‑color custom palette, loops through every worksheet and each chart within, sets each series' fill and border colors from the palette, and saves the updated workbook as a new XLSX file.
class UniformChartColorPalette
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define a uniform custom color palette
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(0x4F81BD), // Blue
                Color.FromArgb(0xC0504D), // Red
                Color.FromArgb(0x9BBB59), // Green
                Color.FromArgb(0x8064A2), // Purple
                Color.FromArgb(0x4BACC6), // Cyan
                Color.FromArgb(0xF79646)  // Orange
            };

            // Apply the custom palette to every series in every chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    int colorIdx = 0;
                    foreach (Series series in chart.NSeries)
                    {
                        // Set the series fill color using the Area's ForegroundColor
                        series.Area.ForegroundColor = customPalette[colorIdx % customPalette.Length];
                        // Optionally set the border color to match
                        series.Border.Color = customPalette[colorIdx % customPalette.Length];
                        colorIdx++;
                    }
                }
            }

            // Save the workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
