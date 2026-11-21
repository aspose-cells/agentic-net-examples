// Title: Resize an Excel slicer by setting explicit width and height with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set slicer width and height | how to programmatically change Excel slicer size using Aspose | adjust slicer dimensions in a workbook with Aspose.Cells .NET | resize slicer to specific points Aspose.Cells example | set slicer layout consistency Aspose.Cells C# code
// Tags: slicer width property Aspose.Cells | slicer height property Aspose.Cells | resize Excel slicer programmatically C# | set slicer dimensions .NET workbook | Aspose.Cells slicer layout adjustment | configure slicer size points Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer class

// The example loads an existing workbook, locates the slicer named "Slicer1" on the first worksheet, assigns a Width of 300 points and a Height of 200 points, ensures the output folder exists, and saves the modified file, demonstrating how to control slicer size for consistent layout using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index/name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the slicer by its name
            Slicer slicer = null;
            try
            {
                slicer = worksheet.Slicers["Slicer1"];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving slicer: {ex.Message}");
            }

            if (slicer == null)
            {
                Console.WriteLine("Error: Slicer \"Slicer1\" not found in the worksheet.");
                return;
            }

            // Assign specific layout dimensions (values are in points)
            slicer.Width = 300;   // Desired width
            slicer.Height = 200;  // Desired height

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the resized slicer
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
