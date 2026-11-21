// Title: Read defined names (named ranges) from an Excel workbook using Aspose.Cells in C#
// Common Searches: how to enumerate named ranges in an Excel file with Aspose.Cells C# | Aspose.Cells C# retrieve all defined names from a workbook | list Excel defined names using Aspose.Cells API in .NET | C# code example for reading named ranges with Aspose.Cells
// Tags: Aspose.Cells read defined names C# | enumerate named ranges Aspose.Cells | C# extract workbook defined names | Aspose.Cells workbook named range iteration

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input file, loads it into an Aspose.Cells Workbook, iterates through workbook.Worksheets.Names, safely obtains each defined name's Name and RefersTo values via reflection, and prints them to the console while handling possible errors.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Input file not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook (no special filter needed for defined names)
            Workbook workbook = new Workbook(filePath);

            // Iterate through all defined names in the workbook
            foreach (var definedName in workbook.Worksheets.Names)
            {
                string nameStr = string.Empty;
                string refersToStr = string.Empty;

                try
                {
                    // Use reflection to obtain the Name and RefersTo properties (covers API variations)
                    var type = definedName.GetType();
                    var nameProp = type.GetProperty("Name");
                    var refersProp = type.GetProperty("RefersTo");

                    nameStr = nameProp?.GetValue(definedName)?.ToString() ?? string.Empty;
                    refersToStr = refersProp?.GetValue(definedName)?.ToString() ?? string.Empty;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read defined name properties: {ex.Message}");
                }

                Console.WriteLine($"Name: {nameStr}, RefersTo: {refersToStr}");
            }

            // (Optional) Save the workbook if modifications are made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
