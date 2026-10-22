// Title: Make an Excel slicer printable using Aspose.Cells for .NET (C#)
// AI Prompts: Set the IsPrintable property of a specific slicer in a loaded workbook with Aspose.Cells C#. | Enable printing for all slicers on a worksheet by iterating the SlicerCollection via Aspose.Cells .NET API. | Update an existing Excel file so its slicer appears in print output, using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# how to set slicer printable property | Make Excel slicer appear when printing using Aspose.Cells .NET | C# example to enable slicer printing in an existing workbook | Set IsPrintable flag on slicer with Aspose.Cells API
// Tags: Aspose.Cells set slicer printable | C# slicer IsPrintable property | Excel slicer print visibility .NET | modify slicer print option Aspose.Cells | load workbook adjust slicer print

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer and SlicerCollection classes

// The code loads an existing workbook, accesses the first worksheet’s slicer collection, sets the slicer’s IsPrintable property to true, and saves the workbook, ensuring the slicer is included when the sheet is printed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the desired worksheet index/name)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of slicers on the worksheet
            SlicerCollection slicers = sheet.Slicers;

            if (slicers.Count > 0)
            {
                // Use the first slicer; alternatively locate by name: slicers["MySlicerName"]
                Slicer slicer = slicers[0];

                // Make the slicer printable so it appears when printing the sheet
                slicer.IsPrintable = true;
            }
            else
            {
                Console.WriteLine("No slicers found on the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
