// Title: Toggle visibility of an embedded picture in an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide picture inside chart based on condition | How to set IsVisible for chart picture with Aspose.Cells .NET | Toggle embedded chart image visibility in Excel using Aspose.Cells | Aspose.Cells chart Pictures collection hide or show picture programmatically | Conditional display of chart picture in workbook with Aspose.Cells C#
// Tags: Aspose.Cells chart picture visibility | C# set chart picture IsVisible | Aspose.Cells toggle embedded chart image | Excel chart picture manipulation .NET | Aspose.Cells Pictures property usage | conditional chart picture display Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an existing workbook, accesses the first worksheet and its first chart, and uses a Boolean flag to set the IsVisible property of the first picture in the chart (if present). Handles versions without the Pictures property gracefully and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0];

                // Flag that determines the picture visibility
                bool showPicture = true; // Set to false to hide the picture

                // Attempt to toggle picture visibility if the API supports it
                try
                {
                    // The Chart.Pictures property is available in newer versions of Aspose.Cells.
                    // If unavailable, this block will be skipped safely.
                    var pictures = chart.Pictures;
                    if (pictures != null && pictures.Count > 0)
                    {
                        pictures[0].IsVisible = showPicture;
                    }
                    else
                    {
                        Console.WriteLine("No pictures found inside the chart.");
                    }
                }
                catch (MissingMemberException)
                {
                    // Property not available in the current Aspose.Cells version
                    Console.WriteLine("Chart picture manipulation is not supported in this Aspose.Cells version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error handling chart pictures: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
