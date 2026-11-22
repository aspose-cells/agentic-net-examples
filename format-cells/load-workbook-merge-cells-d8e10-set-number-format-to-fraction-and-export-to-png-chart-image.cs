// Title: Merge cells D8:E10, apply a fraction number format, and export a chart as PNG using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge D8:E10 and set custom fraction format | Export first worksheet chart to PNG with Aspose.Cells .NET | Apply '# ?/?' number format to a merged range in Excel using Aspose.Cells | Create a sample column chart when none exists and save as PNG with Aspose.Cells | Configure ImageOrPrintOptions for one page per sheet while exporting chart
// Tags: merge cells range Aspose.Cells C# | fraction number format style Aspose.Cells | chart export PNG Aspose.Cells | configure ImageOrPrintOptions for chart export | add column chart programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an existing workbook, merges the range D8:E10 on the first worksheet, applies a custom fraction number format ("# ?/?") to the merged cells, creates a sample column chart if the sheet has none, exports the first chart to a PNG image, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";
        const string chartImageFile = "chart.png";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D8:E10 (row index 7, column index 3, height 3 rows, width 2 columns)
            sheet.Cells.Merge(7, 3, 3, 2);

            // Create a style with fraction number format
            Style fractionStyle = workbook.CreateStyle();
            fractionStyle.Custom = "# ?/?"; // fraction format

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("D8:E10");
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            mergedRange.ApplyStyle(fractionStyle, flag);

            // Prepare image export options for PNG (default format is PNG when file extension is .png)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Export the first chart to PNG if it exists; otherwise create a simple chart and export it
            try
            {
                if (sheet.Charts.Count > 0)
                {
                    Chart existingChart = sheet.Charts[0];
                    existingChart.ToImage(chartImageFile, imgOptions);
                }
                else
                {
                    // Create a sample column chart using data from A1:B5
                    int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart newChart = sheet.Charts[chartIdx];
                    newChart.NSeries.Add("B1:B5", true);
                    newChart.NSeries.CategoryData = "A1:A5";
                    newChart.Title.Text = "Sample Chart";
                    newChart.ToImage(chartImageFile, imgOptions);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chart export error: {ex.Message}");
            }

            // Save the workbook if further changes need to be persisted
            workbook.Save(outputFile);
            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
