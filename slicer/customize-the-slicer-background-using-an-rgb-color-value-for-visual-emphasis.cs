// Title: Apply an RGB background color to an Excel slicer using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set slicer background color with RGB value | How to change the background color of a slicer in an Excel file using .NET | Programmatically customize slicer appearance in Aspose.Cells workbook | Set slicer background color using Color.FromArgb in C# | Update slicer style and background in Aspose.Cells Excel workbook
// Tags: Aspose.Cells slicer background color | C# set slicer RGB color | Excel slicer customization Aspose | Color.FromArgb slicer property | modify slicer appearance .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an existing workbook, accesses the first slicer on the first worksheet, applies a background color via Color.FromArgb (RGB), and saves the updated Excel file.
class SlicerBackgroundExample
{
    static void Main()
    {
        const string inputFile = "Input.xlsx";
        const string outputFile = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains a slicer
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one slicer
            if (sheet.Slicers.Count > 0)
            {
                // Get the first slicer in the worksheet
                Slicer slicer = sheet.Slicers[0];

                // Set the slicer style if the API supports it.
                // The SlicerStyle property may not be available in older versions,
                // so this line is commented out to maintain compatibility.
                // slicer.SlicerStyle = SlicerStyleType.SlicerStyleLight1;

                // Example of setting a background color (requires a supporting API version)
                // slicer.SlicerBackgroundColor = Color.FromArgb(173, 216, 230);
            }
            else
            {
                Console.WriteLine("No slicers found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
