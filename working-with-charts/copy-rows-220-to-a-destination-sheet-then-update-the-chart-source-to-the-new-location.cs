// Title: Copy rows 2‑20 from a source worksheet to a destination worksheet and reassign chart data source using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy rows 2 to 20 from source sheet to destination sheet | How to update chart series formula after moving rows with Aspose.Cells .NET | C# code to copy a range of rows and adjust chart source range in Excel using Aspose.Cells | Aspose.Cells copy rows between worksheets and refresh chart data source | Change chart data source to new worksheet after copying rows in Aspose.Cells C#
// Tags: Aspose.Cells copy rows between worksheets | Aspose.Cells update chart series formula | Aspose.Cells Cells.CopyRows method | Aspose.Cells chart data source adjustment | Aspose.Cells .xlsx workbook manipulation C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, copies rows 2‑20 from the "Source" worksheet to the "Destination" worksheet using Cells.CopyRows, updates any chart on the destination sheet by replacing the source sheet name with the destination sheet name in each series formula, and saves the modified workbook as a new .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get source and destination worksheets (adjust names as needed)
                Worksheet sourceSheet = workbook.Worksheets["Source"];
                Worksheet destinationSheet = workbook.Worksheets["Destination"];

                if (sourceSheet == null || destinationSheet == null)
                {
                    Console.WriteLine("Source or Destination worksheet not found.");
                    return;
                }

                // Copy rows 2‑20 (zero‑based index 1‑19) from source to destination,
                // starting at row 2 in the destination sheet
                int startRow = 1;               // row 2 (zero‑based)
                int endRow = 19;                // row 20 (zero‑based)
                int destinationStartRow = 1;    // row 2 in destination sheet
                int rowsToCopy = endRow - startRow + 1;

                destinationSheet.Cells.CopyRows(sourceSheet.Cells, startRow, destinationStartRow, rowsToCopy);

                // Update chart source to point to the new location (if a chart exists)
                if (destinationSheet.Charts.Count > 0)
                {
                    try
                    {
                        Chart chart = destinationSheet.Charts[0];

                        foreach (Series series in chart.NSeries)
                        {
                            // Use reflection to access the Formula property (covers API variations)
                            PropertyInfo formulaProp = series.GetType().GetProperty("Formula", BindingFlags.Public | BindingFlags.Instance);
                            if (formulaProp != null && formulaProp.CanRead && formulaProp.CanWrite)
                            {
                                string originalFormula = formulaProp.GetValue(series) as string ?? string.Empty;
                                string updatedFormula = originalFormula.Replace(sourceSheet.Name, destinationSheet.Name);
                                formulaProp.SetValue(series, updatedFormula);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error updating chart series: {ex.Message}");
                    }
                }

                // Ensure the output directory exists (handle case where outputPath has no directory)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
