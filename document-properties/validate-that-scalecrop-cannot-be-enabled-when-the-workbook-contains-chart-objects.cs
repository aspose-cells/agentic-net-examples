// Title: Validate that ScaleCrop is disabled for all pictures when an Excel workbook contains chart objects using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans each worksheet for charts and throws an InvalidOperationException if any picture has IsScaleCrop set to true. | Create a C# validation method that prevents enabling the ScaleCrop property on images when the workbook includes at least one chart, using reflection to support older Aspose.Cells versions.
// Common Searches: Aspose.Cells C# check workbook for charts before allowing picture ScaleCrop | how to detect chart objects and enforce IsScaleCrop false on images in Excel with Aspose.Cells | C# validate picture ScaleCrop property when Excel file contains chart objects Aspose.Cells | throw exception for enabled ScaleCrop on pictures in a charted workbook using Aspose.Cells .NET
// Tags: ScaleCrop property validation with Aspose.Cells | chart presence check for picture ScaleCrop | C# workbook chart detection Aspose.Cells | picture IsScaleCrop verification in Excel .NET | exception handling for invalid ScaleCrop setting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, determines if any worksheet contains chart objects, then iterates through all pictures to ensure the IsScaleCrop property is false. If a picture has ScaleCrop enabled while charts exist, an InvalidOperationException is thrown; otherwise the workbook is saved.
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
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine if any worksheet contains a chart object
            bool hasChart = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    hasChart = true;
                    break;
                }
            }

            // If charts exist, ensure no picture has ScaleCrop enabled
            if (hasChart)
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Picture picture in sheet.Pictures)
                    {
                        // Use reflection to check for the IsScaleCrop property (may not exist in older versions)
                        var propInfo = picture.GetType().GetProperty("IsScaleCrop");
                        if (propInfo != null && propInfo.PropertyType == typeof(bool))
                        {
                            bool isScaleCrop = (bool)propInfo.GetValue(picture);
                            if (isScaleCrop)
                            {
                                throw new InvalidOperationException(
                                    $"ScaleCrop is enabled on picture '{picture.Name}' in sheet '{sheet.Name}' while the workbook contains chart objects.");
                            }
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
