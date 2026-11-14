// Title: How to disable printing for every slicer in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set slicer IsPrintable false for all worksheets | remove slicers from print area using Aspose.Cells .NET | iterate over slicer collection and disable printing in C# | programmatically hide slicers from printing in Excel with Aspose.Cells | C# code to make slicers non‑printable in an existing .xlsx file
// Tags: Aspose.Cells slicer IsPrintable property | disable slicer printing Excel .NET | iterate slicer collection workbook C# | set slicer non‑printable Aspose.Cells | modify slicer printable flag programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer and SlicerCollection

// Loads an existing .xlsx file, loops through each worksheet's SlicerCollection, sets each slicer's IsPrintable property to false, and saves the workbook, ensuring slicers are excluded from print output.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the slicer collection of the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Set each slicer's printable flag to false
                foreach (Slicer slicer in slicers)
                {
                    slicer.IsPrintable = false;
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
