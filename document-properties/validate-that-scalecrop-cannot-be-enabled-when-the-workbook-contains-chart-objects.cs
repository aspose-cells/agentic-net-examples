// Title: Validate that ScaleCrop is disabled on pictures when an Excel workbook contains chart objects using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if any picture has ScaleCrop enabled when workbook contains charts | how to read the ScaleCrop property of an image in an Excel file with Aspose.Cells using reflection | throw exception if ScaleCrop is true on a picture and the workbook has chart objects in Aspose.Cells
// Tags: Aspose.Cells picture ScaleCrop validation | chart detection before image processing Aspose.Cells | C# reflection read ScaleCrop property Aspose.Cells | prevent ScaleCrop on pictures in chart‑containing workbook | Excel workbook chart presence check Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook with Aspose.Cells, checks whether any worksheet contains chart objects, and if so iterates through all shapes to find picture objects. Using reflection it reads the boolean ScaleCrop property (if available) and throws an InvalidOperationException when ScaleCrop is enabled on any picture, then saves the workbook unchanged.
class ScaleCropValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file '{inputPath}' not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            bool containsChart = false;

            // Determine if any worksheet contains chart objects
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    containsChart = true;
                    break;
                }
            }

            // If charts exist, ensure no picture has ScaleCrop enabled
            if (containsChart)
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        if (shape is Picture picture)
                        {
                            // Use reflection to read ScaleCrop if the property exists in the current version
                            bool scaleCropEnabled = false;
                            PropertyInfo prop = picture.GetType().GetProperty("ScaleCrop", BindingFlags.Public | BindingFlags.Instance);
                            if (prop != null && prop.PropertyType == typeof(bool))
                            {
                                scaleCropEnabled = (bool)prop.GetValue(picture);
                            }

                            if (scaleCropEnabled)
                            {
                                throw new InvalidOperationException(
                                    $"ScaleCrop is enabled on picture '{picture.Name}' in sheet '{sheet.Name}' while the workbook contains chart objects.");
                            }
                        }
                    }
                }
            }

            // Save the workbook (unchanged)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
