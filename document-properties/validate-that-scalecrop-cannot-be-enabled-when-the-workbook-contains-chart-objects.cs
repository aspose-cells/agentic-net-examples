// Title: Validate that picture ScaleCrop cannot be enabled when an Excel workbook contains chart objects using Aspose.Cells for .NET
// Common Searches: how to prevent picture ScaleCrop when workbook has charts using Aspose.Cells C# | Aspose.Cells C# verify no ScaleCrop on images if worksheet contains a chart | C# code to detect charts before setting picture ScaleCrop with Aspose.Cells | validate picture ScaleCrop restriction with chart objects in Aspose.Cells .NET | Aspose.Cells example to check for charts before applying ScaleCrop to pictures
// Tags: Aspose.Cells picture ScaleCrop validation | detect charts in Excel workbook Aspose.Cells | prevent ScaleCrop on images when charts present | C# Aspose.Cells workbook chart detection | Excel picture processing with chart constraints

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an Excel file with Aspose.Cells, scans all worksheets to see if any chart objects exist, and if charts are found iterates through each picture as a placeholder for future ScaleCrop validation (the current API does not expose a ScaleCrop property). It then saves the workbook while handling I/O and runtime errors.
class Program
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine if the workbook contains any chart objects
            bool containsChart = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    containsChart = true;
                    break;
                }
            }

            // If charts are present, validate that ScaleCrop is not enabled on any picture
            // Note: Aspose.Cells' Picture class does not expose a ScaleCrop property in the current version.
            // Therefore, this validation is omitted. Adjust as needed for newer versions that support ScaleCrop.
            if (containsChart)
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Picture picture in sheet.Pictures)
                    {
                        // Placeholder for future ScaleCrop validation if the API provides it.
                        // Example: if (picture.ScaleCrop) { throw new InvalidOperationException(...); }
                    }
                }
            }

            // Save the workbook (optional)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
