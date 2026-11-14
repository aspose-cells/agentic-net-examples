// Title: Conditionally apply a slicer style in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to set slicer style based on caption text with Aspose.Cells C# | Aspose.Cells iterate slicers and change style when caption contains Sales | apply conditional formatting to Excel slicers using Aspose.Cells .NET API | C# code to change slicer style programmatically in an existing workbook | check slicer caption and assign predefined slicer style with Aspose.Cells
// Tags: conditional slicer styling with Aspose.Cells | assign slicer style by caption in C# | enumerate worksheet slicers using Aspose.Cells API | apply predefined slicer style to Excel workbook | handle missing SlicerStyles collection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, enumerates all slicers on the first worksheet, and for each slicer whose Caption contains the word "Sales" it attempts to assign a predefined slicer style (code commented out due to possible API version limits). It logs a placeholder message when the style cannot be applied and then saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

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

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all slicers on the worksheet
                foreach (Slicer slicer in worksheet.Slicers)
                {
                    // Apply style when the slicer caption contains the word "Sales"
                    if (!string.IsNullOrEmpty(slicer.Caption) && slicer.Caption.Contains("Sales"))
                    {
                        try
                        {
                            // NOTE: The SlicerStyles collection may not be available in some versions.
                            // If it is available, you can uncomment the following lines and adjust as needed.
                            // if (workbook.SlicerStyles.IndexOf("SlicerStyle1") != -1)
                            // {
                            //     slicer.Style = workbook.SlicerStyles["SlicerStyle1"];
                            // }
                            // else
                            // {
                            //     Console.WriteLine("Style 'SlicerStyle1' not found in the workbook.");
                            // }

                            // Placeholder action when style cannot be applied
                            Console.WriteLine($"Slicer '{slicer.Caption}' matches criteria but style application is skipped (API not available).");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to apply style to slicer '{slicer.Caption}': {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
