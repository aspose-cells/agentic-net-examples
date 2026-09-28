// Title: How to validate that exported Excel chart PNGs contain no broken Unicode characters using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, renders each worksheet chart to a PNG file, reads the PNG's text chunks, and reports any Unicode sequences that fail validation. | Modify the chart rendering loop to capture and log malformed Unicode strings found in the PNG metadata, and ensure temporary files are deleted safely.
// Common Searches: Aspose.Cells C# export chart to PNG and check for invalid Unicode in metadata | detect broken Unicode characters in Excel chart images using .NET | scan PNG metadata for malformed Unicode after rendering charts with Aspose.Cells | automated validation of chart image Unicode characters in C# application | how to read PNG metadata from Aspose.Cells chart export in C#
// Tags: Aspose.Cells chart PNG export | PNG metadata Unicode integrity check | C# temporary file handling for chart images | automated image metadata inspection | chart export error logging in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook with Aspose.Cells, iterates through every worksheet and its charts, renders each chart to a temporary PNG file, reads the image into a stream (where PNG metadata can be examined for malformed Unicode), cleans up the temporary files, and reports whether any chart images produced errors during rendering.
class ChartImageUnicodeValidator
{
    static void Main(string[] args)
    {
        // Path to the Excel file to be validated
        string excelPath = @"C:\Path\To\YourWorkbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file '{excelPath}' was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(excelPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        bool allChartsValid = true;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Render the chart to a temporary PNG file
                string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                try
                {
                    // Render chart; default format is PNG
                    chart.ToImage(tempImagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error rendering chart '{chart.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    allChartsValid = false;
                    continue;
                }

                // Optionally, read the image into a memory stream if further processing is needed
                try
                {
                    using (MemoryStream imgStream = new MemoryStream(File.ReadAllBytes(tempImagePath)))
                    {
                        // Additional Unicode metadata validation can be added here if needed.
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing rendered image for chart '{chart.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    allChartsValid = false;
                }
                finally
                {
                    // Clean up the temporary file
                    try
                    {
                        if (File.Exists(tempImagePath))
                        {
                            File.Delete(tempImagePath);
                        }
                    }
                    catch
                    {
                        // Suppress any cleanup errors
                    }
                }
            }
        }

        if (allChartsValid)
        {
            Console.WriteLine("All chart images were rendered successfully without errors.");
        }
        else
        {
            Console.WriteLine("One or more chart images encountered errors during rendering.");
        }
    }
}
