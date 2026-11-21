// Title: List all named ranges in an Excel workbook and print their RefersTo addresses with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to get all defined names from an Excel file | how to print named range addresses from a workbook using Aspose.Cells | enumerate worksheet named ranges and display RefersTo values in .NET | Aspose.Cells retrieve NameCollection items and output to console | list Excel named ranges programmatically with Aspose.Cells C# code
// Tags: enumerate named ranges Aspose.Cells | retrieve RefersTo address C# | Aspose.Cells NameCollection iteration | list defined names in Excel workbook | C# console output named range addresses

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample checks for the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, accesses the Worksheets.Names collection, iterates through each Name object, and writes the defined name together with its RefersTo address (e.g., =Sheet1!$A$1:$B$2) to the console, handling any exceptions that may occur.
    class Program
    {
        static void Main()
        {
            try
            {
                string filePath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Get the collection of all named ranges in the workbook
                NameCollection namedRanges = workbook.Worksheets.Names;

                // Enumerate each named range and output its address
                foreach (Name namedRange in namedRanges)
                {
                    // namedRange.RefersTo contains the address (e.g., =Sheet1!$A$1:$B$2)
                    Console.WriteLine($"{namedRange.Text}: {namedRange.RefersTo}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
