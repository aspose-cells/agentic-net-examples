// Title: Remove a named slicer from an Excel workbook and save the file using Aspose.Cells for .NET
// Common Searches: how to delete a slicer named 'Slicer1' in an Excel file with Aspose.Cells C# | Aspose.Cells C# remove slicer from worksheet and save workbook | C# code to find and remove slicer by name using Aspose.Cells | programmatically delete Excel slicer and clean up connections with Aspose.Cells | remove slicer from Excel workbook using Aspose.Cells .NET example
// Tags: Aspose.Cells slicer deletion C# | C# remove Excel slicer programmatically | Aspose.Cells workbook slicer cleanup | Excel slicer removal using Aspose.Cells | Aspose.Cells remove slicer by name

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, searches each worksheet for a slicer named "Slicer1", removes the slicer when found, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the slicer to be removed
            string slicerName = "Slicer1";

            // Search all worksheets for the slicer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int slicerIndex = -1;

                // Locate slicer by name within the worksheet
                for (int i = 0; i < sheet.Slicers.Count; i++)
                {
                    if (sheet.Slicers[i].Name == slicerName)
                    {
                        slicerIndex = i;
                        break;
                    }
                }

                if (slicerIndex != -1)
                {
                    // Remove the slicer from the worksheet
                    sheet.Slicers.RemoveAt(slicerIndex);

                    // Slicer found and removed; exit the loop
                    break;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or display the exception details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
