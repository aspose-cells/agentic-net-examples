// Title: Merge cells D8:E10, apply a fraction number format, and export the first worksheet chart to PNG using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge cells D8:E10 and set custom fraction format | How to export an Excel chart to PNG with Aspose.Cells .NET | Apply '# ?/?' number format to a merged range using Aspose.Cells in C# | Save workbook after merging cells and rendering chart image with Aspose.Cells | Create fraction style and apply to merged cells Aspose.Cells example
// Tags: merge cells D8:E10 Aspose.Cells C# | fraction custom number format Aspose.Cells | export chart to PNG Aspose.Cells .NET | apply style to merged range Aspose.Cells | chart rendering image Aspose.Cells | save workbook after cell merge Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The C# program loads an existing Excel workbook, merges the range D8:E10 on the first worksheet, creates a custom fraction number format "# ?/?" and applies it to the merged cells, checks for a chart on the sheet and, if present, renders the first chart to a PNG image, and finally saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string chartImagePath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D8:E10 (zero‑based indices: row 7, column 3)
            sheet.Cells.Merge(7, 3, 3, 2); // firstRow, firstColumn, totalRows, totalColumns

            // Create a fraction number format style
            Style fractionStyle = workbook.CreateStyle();
            fractionStyle.Custom = "# ?/?";

            // Apply the style to the merged range
            AsposeRange mergedRange = sheet.Cells.CreateRange("D8:E10");
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            mergedRange.ApplyStyle(fractionStyle, flag);

            // Export the first chart (if any) to a PNG image
            if (sheet.Charts.Count > 0)
            {
                try
                {
                    Chart chart = sheet.Charts[0];
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions(); // defaults to PNG
                    chart.ToImage(chartImagePath, imgOptions);
                    Console.WriteLine($"Chart exported to '{chartImagePath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
