// Title: C# Aspose.Cells example: List slicer names from an Excel worksheet and save them to a text file
// Common Searches: how to enumerate slicer objects in an Excel file using Aspose.Cells C# | export Excel slicer names to a .txt file with Aspose.Cells for .NET | retrieve all slicer names from the first worksheet in a workbook using C# | Aspose.Cells C# code to write slicer collection to a text document | list slicer names in a spreadsheet and save to file programmatically
// Tags: slicer collection iteration Aspose.Cells | export slicer identifiers to plain text | worksheet slicer enumeration C# | Aspose.Cells save slicer list to file | extract slicer names from Excel workbook | C# write slicer names to .txt

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The sample loads an existing Excel workbook with Aspose.Cells, accesses the first worksheet's SlicerCollection, iterates through each Slicer to obtain its Name property, and writes each name line‑by‑line to a specified text file, handling missing files and runtime exceptions.
class ListSlicerNames
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "slicer_names.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the slicer collection on the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Write slicer names to a text file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (Slicer slicer in slicers)
                {
                    writer.WriteLine(slicer.Name);
                }
            }

            Console.WriteLine($"Slicer names have been written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
