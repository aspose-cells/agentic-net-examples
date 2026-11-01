// Title: Batch update linked shape cell references and chart series data sources across all worksheets using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells update linked shape cell reference across multiple worksheets | how to replace chart series range addresses in all worksheets with Aspose.Cells | bulk modify shape links and chart data sources in an Excel file using Aspose.Cells .NET | programmatically change linked shape targets in a workbook with Aspose.Cells | iterate through worksheets to update chart series values using Aspose.Cells C#
// Tags: linked shape cell reference update Aspose.Cells | chart series data source replacement Aspose.Cells | bulk modify shapes across worksheets .NET | batch update chart ranges Aspose.Cells | iterate worksheets to change linked shapes C# | dynamic chart series manipulation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an existing workbook (or creates a new one), defines a dictionary that maps old cell/range addresses to new ones, iterates over every worksheet, optionally updates linked shapes' cell references, and updates each chart's series Values, XValues, and CategoryData using dynamic objects. Finally, it saves the modified workbook to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load \"{inputPath}\": {loadEx.Message}");
                    Console.WriteLine("Creating a new workbook instead.");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Mapping of old linked cell/range addresses to new ones
            var addressMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Example mappings:
                // { "Sheet1!A1", "Sheet1!B1" },
                // { "Sheet2!C5:D10", "Sheet2!E5:F10" }
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // ----- Update linked shapes (if supported) -----
                /*
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsLinkedToCell)
                    {
                        string oldLink = shape.LinkToCell;
                        if (!string.IsNullOrEmpty(oldLink) && addressMap.TryGetValue(oldLink, out string newLink))
                        {
                            shape.LinkToCell = newLink;
                        }
                    }
                }
                */

                // ----- Update chart series data sources -----
                foreach (Chart chart in sheet.Charts)
                {
                    // Use dynamic to avoid compile‑time dependency on ChartSeries type
                    foreach (var seriesObj in chart.NSeries)
                    {
                        dynamic series = seriesObj;

                        // Update Values range
                        string oldValues = series.Values as string;
                        if (!string.IsNullOrEmpty(oldValues) && addressMap.TryGetValue(oldValues, out string newValues))
                        {
                            series.Values = newValues;
                        }

                        // Update XValues range
                        string oldXValues = series.XValues as string;
                        if (!string.IsNullOrEmpty(oldXValues) && addressMap.TryGetValue(oldXValues, out string newXValues))
                        {
                            series.XValues = newXValues;
                        }

                        // Update CategoryData range
                        string oldCategory = series.CategoryData as string;
                        if (!string.IsNullOrEmpty(oldCategory) && addressMap.TryGetValue(oldCategory, out string newCategory))
                        {
                            series.CategoryData = newCategory;
                        }
                    }
                }
            }

            // Ensure the directory for the output file exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
