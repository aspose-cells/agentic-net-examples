// Title: Extract all charts from an Excel workbook and save each as a PNG file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, loops through every worksheet and its charts, and writes each chart to a PNG image with a unique, sanitized filename. | Create a reusable method that accepts a workbook path and an output directory, configures ImageOrPrintOptions for PNG, ensures the directory exists, and exports all embedded charts to separate PNG files.
// Common Searches: asp.net export all charts from an Excel file to PNG using Aspose.Cells | c# extract chart images from .xlsx workbook with Aspose.Cells | how to save each chart in a workbook as individual PNG files in C# | Aspose.Cells chart ToImage example for multiple worksheets | C# loop through worksheets and charts to generate PNG images
// Tags: export chart to png Aspose.Cells | iterate worksheets and charts C# | chart image rendering options Aspose.Cells | sanitize filename for chart export C# | create output directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program loads a specified Excel workbook, verifies its presence, iterates through every worksheet and each chart it contains, configures PNG rendering options, sanitizes generated filenames, creates necessary output folders, and saves each chart as an individual PNG file while logging successes and handling errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing charts
            Workbook workbook = new Workbook(inputPath);
            int chartIndex = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Set image options (default format is PNG)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            HorizontalResolution = 96,
                            VerticalResolution = 96
                        };

                        // Build a unique file name for the chart image
                        string fileName = $"Chart_{chartIndex}_{sheet.Name}_{chart.Name}.png";

                        // Replace any invalid filename characters
                        foreach (char invalidChar in Path.GetInvalidFileNameChars())
                        {
                            fileName = fileName.Replace(invalidChar, '_');
                        }

                        // Ensure the directory for the output file exists
                        string directory = Path.GetDirectoryName(fileName);
                        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        // Save the chart as a PNG file
                        using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                        {
                            chart.ToImage(stream, imgOptions);
                        }

                        Console.WriteLine($"Saved chart image: {fileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart '{chart.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }

                    chartIndex++;
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
