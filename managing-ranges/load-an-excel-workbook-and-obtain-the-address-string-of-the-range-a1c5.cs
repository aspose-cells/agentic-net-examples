// Title: Load an Excel workbook and retrieve the address string of range A1:C5 with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook and get range address A1:C5 | How to use Aspose.Cells CreateRange to obtain RefersTo string in .NET | Reading the address of a specific cell range from an Excel file using Aspose.Cells | Aspose.Cells example for extracting range address without opening Excel | Retrieve Excel range address programmatically with Aspose.Cells C#
// Tags: Aspose.Cells load workbook | Aspose.Cells create range A1:C5 | Aspose.Cells RefersTo property | C# extract Excel range address | Aspose.Cells range address string

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, creates a range covering cells A1:C5 on the first worksheet, obtains the range's address via the RefersTo property, and prints the address to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create a range that refers to cells A1:C5 in the first worksheet
            // Use the fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange("A1:C5");

            // Obtain the address string of the range (e.g., "A1:C5")
            string address = range.RefersTo;

            // Output the address
            Console.WriteLine(address);
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
