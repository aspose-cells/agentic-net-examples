// Title: C# Example: Retrieve Chart Shape Reflection Blur Using Aspose.Cells for .NET
// Description: Loads an existing or new workbook, walks through each worksheet and its shapes, filters ChartShape objects, reads the Reflection.Blur property (default 0 if missing), stores the blur radius in a dictionary keyed by worksheet and shape name, outputs the results, and saves the workbook unchanged.
// Keywords: Aspose.Cells | C# | .NET | chart shape reflection blur | Reflection.Blur property | retrieve chart effects | Excel shape reflection | dictionary of blur values | Aspose.Cells example | GitHub source code
// Common Searches: Aspose.Cells get chart reflection blur C# | read reflection effect of chart shapes .NET | extract blur radius from Excel chart shape | store chart reflection values in dictionary | Aspose.Cells chart shape reflection property example
// Developer Intent: Read the blur radius of the reflection effect for every chart shape in a workbook and collect the values in a dictionary.
// Use Cases: Generate a design audit that lists reflection blur settings for all charts across worksheets. | Validate visual consistency of chart reflections before publishing an Excel file. | Compare or adjust chart reflection blur values programmatically during automated report generation.
// AI Prompts: Write C# code with Aspose.Cells to set the reflection blur of all chart shapes to a given value. | Create a method that returns a dictionary mapping chart identifiers to their reflection blur values for a workbook. | Explain how to handle null Reflection objects safely when extracting blur values from chart shapes in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing or new workbook, walks through each worksheet and its shapes, filters ChartShape objects, reads the Reflection.Blur property (default 0 if missing), stores the blur radius in a dictionary keyed by worksheet and shape name, outputs the results, and saves the workbook unchanged.
class RetrieveChartReflectionBlur
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Store blur values for each chart shape.
            var chartBlurValues = new Dictionary<string, double>();

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet.
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only chart shapes using pattern matching.
                    if (shape is ChartShape chartShape && chartShape.Chart != null)
                    {
                        try
                        {
                            // Get the reflection effect of the chart shape.
                            ReflectionEffect reflection = chartShape.Reflection;

                            // Use worksheet name and shape name as a unique key.
                            string key = $"{sheet.Name}_{chartShape.Name}";

                            // Store the blur radius (in points). If reflection is null, default to 0.
                            chartBlurValues[key] = reflection?.Blur ?? 0;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to retrieve reflection for chart \"{chartShape.Name}\": {ex.Message}");
                        }
                    }
                }
            }

            // Output collected blur values.
            foreach (var kvp in chartBlurValues)
            {
                Console.WriteLine($"Chart: {kvp.Key}, Reflection Blur: {kvp.Value}");
            }

            // Save the workbook (no modifications made in this example).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
