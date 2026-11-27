// Title: Duplicating a chart and assigning a new category axis range in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# duplicate existing chart and change its category axis range | Copy a chart in an Excel file and set a new category data range using Aspose.Cells | How to clone a chart and modify the category axis values in .NET with Aspose.Cells | C# Aspose.Cells example for copying chart series and updating category range | Programmatically copy an Excel chart and assign a different category axis range with Aspose.Cells
// Tags: chart duplication Aspose.Cells C# | set category axis range Aspose.Cells | copy chart series Aspose.Cells | excel chart cloning Aspose.Cells .NET | modify chart data range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DuplicateChartExample
{
    // The sample loads an existing workbook, duplicates the first chart on the first worksheet, copies each series and its properties, assigns a new category axis range (C2:C5) to the duplicated chart, replicates the title and legend settings, and saves the result as a new Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart to duplicate.
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("The worksheet does not contain any charts to duplicate.");
                    return;
                }

                // Source chart (first chart in the worksheet).
                Chart sourceChart = worksheet.Charts[0];

                // Add a new chart. Position is set to a default area; adjust as needed.
                int newChartIndex = worksheet.Charts.Add(
                    sourceChart.Type,
                    0,   // top row
                    0,   // left column
                    15,  // bottom row
                    5);  // right column

                Chart duplicatedChart = worksheet.Charts[newChartIndex];

                // Clear any automatically added series.
                duplicatedChart.NSeries.Clear();

                // Duplicate each series from the source chart.
                foreach (Series srcSeries in sourceChart.NSeries)
                {
                    try
                    {
                        // Add series using the same data range; assume vertical orientation.
                        duplicatedChart.NSeries.Add(srcSeries.Values, true);

                        // Retrieve the newly added series (last in the collection).
                        Series dupSeries = duplicatedChart.NSeries[duplicatedChart.NSeries.Count - 1];

                        // Copy common properties.
                        dupSeries.Name = srcSeries.Name;
                        dupSeries.Type = srcSeries.Type;
                        dupSeries.Marker.MarkerStyle = srcSeries.Marker.MarkerStyle;
                    }
                    catch (Exception exSeries)
                    {
                        Console.WriteLine($"Failed to duplicate a series: {exSeries.Message}");
                    }
                }

                // Assign a new category axis range to the duplicated chart.
                string newCategoryRange = "C2:C5"; // Adjust as required.
                duplicatedChart.NSeries.CategoryData = newCategoryRange;

                // Copy title if it exists.
                if (sourceChart.Title != null && !string.IsNullOrEmpty(sourceChart.Title.Text))
                {
                    duplicatedChart.Title.Text = sourceChart.Title.Text + " (Copy)";
                }

                // Copy legend position.
                duplicatedChart.Legend.Position = sourceChart.Legend.Position;

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
