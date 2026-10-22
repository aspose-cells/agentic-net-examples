// Title: How to enumerate and print all slicer names from the first worksheet of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, verifies the file exists, accesses the first worksheet, and writes each slicer's Name to the console. | Create a reusable C# method that takes a Worksheet object and returns a List<string> of all slicer names on that sheet using Aspose.Cells. | Demonstrate exception handling while iterating the SlicerCollection of a worksheet in Aspose.Cells, covering file‑not‑found and runtime errors.
// Common Searches: aspnet aspocells get slicer names from worksheet c# | c# enumerate slicers in excel file using Aspose.Cells | how to list all slicer objects in a workbook with Aspose.Cells .NET | retrieve slicer collection from first worksheet Aspose.Cells example | Aspose.Cells read slicer properties from .xlsx in C#
// Tags: enumerate slicer collection Aspose.Cells | list slicer names C# | worksheet slicer iteration .NET | Aspose.Cells read slicer properties | Excel slicer enumeration using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer and SlicerCollection

// // Loads an Excel workbook, accesses the first worksheet, obtains its SlicerCollection, and writes each slicer's Name to the console. Includes file existence check and comprehensive exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists to prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the collection of slicers on the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through each slicer and output its name
            foreach (Slicer slicer in slicers)
            {
                Console.WriteLine(slicer.Name);
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
