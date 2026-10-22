// Title: Load an Excel workbook and get the address string of range A1:C5 with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to retrieve the address of a specific range after loading a workbook | Get range reference string A1:C5 using Aspose.Cells .NET | Aspose.Cells create range and read RefersTo property in C# example | Load Excel file and obtain range address without using Excel Interop | Aspose.Cells read cell range address from first worksheet in C#
// Tags: Aspose.Cells load workbook C# | Aspose.Cells create range A1:C5 | Aspose.Cells RefersTo property usage | Aspose.Cells get range address string | Aspose.Cells first worksheet range address

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of "input.xlsx", loads it into an Aspose.Cells Workbook, accesses the first worksheet, creates the range A1:C5, reads the range's RefersTo property to obtain its address string, prints the address, and includes basic exception handling.
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

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create the range A1:C5 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "C5");

            // Obtain the address string of the range
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
