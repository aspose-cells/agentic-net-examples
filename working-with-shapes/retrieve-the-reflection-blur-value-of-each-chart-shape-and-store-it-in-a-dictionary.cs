// Title: How to read the reflection blur value of each chart shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get reflection blur from chart shapes in a workbook | Read chart shape reflection effect blur property using Aspose.Cells for .NET | Extract reflection blur values of Excel chart objects with Aspose.Cells C# | Iterate worksheet shapes and retrieve reflection blur of ChartShape in Aspose.Cells | Store chart shape reflection blur values in a dictionary using Aspose.Cells
// Tags: chartshape reflection blur extraction | worksheet shapes iteration Aspose.Cells | reflection effect property reading | dictionary storage of shape blur values | excel chart shape properties Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartReflectionBlur
{
    // The example loads an Excel workbook with Aspose.Cells, loops through all worksheets and their shapes, identifies ChartShape objects, reads each shape's Reflection?.Blur value (defaulting to 0.0 when unavailable), and saves the blur values in a Dictionary keyed by "WorksheetName:ShapeName" before outputting the results.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the input Excel file (required).
        /// </param>
        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args == null || args.Length < 1)
            {
                Console.WriteLine("Usage: AsposeCellsChartReflectionBlur.exe <input_excel_file_path>");
                return;
            }

            string inputPath = args[0];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Dictionary to hold the blur values.
                // Key format: "WorksheetName:ShapeName"
                Dictionary<string, double> chartReflectionBlur = new Dictionary<string, double>();

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Identify chart shapes (ChartShape derives from Shape)
                        if (shape is ChartShape)
                        {
                            double blur = 0.0;
                            try
                            {
                                // Access the ReflectionEffect; it is always available via the property.
                                blur = shape.Reflection?.Blur ?? 0.0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Unable to read reflection blur for shape \"{shape.Name}\". {ex.Message}");
                            }

                            // Build a readable key
                            string shapeIdentifier = !string.IsNullOrEmpty(shape.Name)
                                ? shape.Name
                                : "UnnamedShape";

                            string dictKey = $"{worksheet.Name}:{shapeIdentifier}";

                            // Store the blur value
                            chartReflectionBlur[dictKey] = blur;
                        }
                    }
                }

                // Output the collected values
                Console.WriteLine("Reflection Blur values for chart shapes:");
                if (chartReflectionBlur.Count == 0)
                {
                    Console.WriteLine("No chart shapes were found in the workbook.");
                }
                else
                {
                    foreach (KeyValuePair<string, double> entry in chartReflectionBlur)
                    {
                        Console.WriteLine($"{entry.Key} => Blur = {entry.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Report any unexpected errors without swallowing them
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
