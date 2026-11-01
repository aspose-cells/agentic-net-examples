// Title: Switch an Aspose.Cells worksheet to Normal view to hide page‑break outlines in C#
// AI Prompts: Generate C# code that sets a worksheet's ViewOptions.ShowPageBreakOutline to false using Aspose.Cells. | Show how to change a worksheet to Normal view mode with Aspose.Cells for .NET so page‑break outlines are not displayed. | Write a C# example that uses reflection to disable the ShowPageBreakOutline property when the ViewOptions API may be unavailable.
// Common Searches: Aspose.Cells C# change worksheet view to normal hide page break outlines | disable page break outline in Excel file using Aspose.Cells .NET | set ShowPageBreakOutline false programmatically Aspose.Cells | how to remove page break outlines from a workbook with Aspose.Cells C# | use reflection to access ViewOptions in Aspose.Cells when property missing
// Tags: worksheet view mode normal Aspose.Cells | hide page break outlines Aspose.Cells | ShowPageBreakOutline property C# | reflection access ViewOptions Aspose.Cells | modify worksheet view options .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing workbook, accesses the first worksheet, disables page‑break outlines by setting ViewOptions.ShowPageBreakOutline to false (using reflection for version compatibility), and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Temp\Input.xlsx";
            string outputPath = @"C:\Temp\Output.xlsx";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Attempt to hide page‑break outlines using reflection.
                // This avoids compile‑time errors if the ViewOptions API is unavailable in the referenced version.
                var viewOptionsProp = sheet.GetType().GetProperty("ViewOptions");
                if (viewOptionsProp != null)
                {
                    var viewOptions = viewOptionsProp.GetValue(sheet);
                    var showOutlineProp = viewOptions?.GetType().GetProperty("ShowPageBreakOutline");
                    if (showOutlineProp != null && showOutlineProp.CanWrite)
                    {
                        showOutlineProp.SetValue(viewOptions, false);
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
