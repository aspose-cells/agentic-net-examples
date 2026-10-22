// Title: How to list all slicer names from an Excel worksheet and save them to a text file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, iterates over the worksheet's Slicers collection, and writes each slicer’s Name to a .txt file. | Provide a robust Aspose.Cells example that verifies the input file, extracts slicer names from the first sheet, and outputs them to a plain‑text file with proper error handling.
// Common Searches: C# Aspose.Cells get slicer names from first worksheet | save Excel slicer list to text file using Aspose.Cells | how to iterate worksheet.Slicers collection in Aspose.Cells C# | export slicer names to .txt with Aspose.Cells .NET | Aspose.Cells example for extracting slicer names
// Tags: Aspose.Cells slicer name extraction | C# export slicer list to text | worksheet.Slicers iteration Aspose.Cells | export slicer collection to plain text | Aspose.Cells .NET slicer enumeration

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, loops through its Slicers collection, and writes each slicer's Name to 'slicer_names.txt'. It includes file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputFile = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the output text file path
            string outputFile = "slicer_names.txt";

            // Write each slicer's name to the text file
            using (StreamWriter writer = new StreamWriter(outputFile))
            {
                // Iterate through the slicers collection
                for (int i = 0; i < worksheet.Slicers.Count; i++)
                {
                    var slicer = worksheet.Slicers[i];
                    writer.WriteLine(slicer.Name);
                }
            }

            Console.WriteLine($"Slicer names have been written to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
