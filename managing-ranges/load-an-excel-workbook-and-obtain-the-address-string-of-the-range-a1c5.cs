// Title: C# – Load an Excel workbook and get the address of range A1:C5 with Aspose.Cells
// Description: A concise example that checks for an existing Excel file, creates a new workbook if needed, loads the first worksheet, defines the range A1:C5 using Aspose.Cells.Range, reads the Range.Address property, and prints the address string to the console.
// Keywords: Aspose.Cells | .NET | C# | load workbook | range address | A1:C5 | CreateRange | Range.Address | Excel automation
// Common Searches: Aspose.Cells get range address C# | How to retrieve A1:C5 address using Aspose.Cells | C# Aspose.Cells create range and read address | Load workbook and obtain range string Aspose.Cells .NET | Aspose.Cells Range.Address property example
// Developer Intent: Load an Excel file and obtain the textual address of the cell block A1:C5.
// Use Cases: Log the range address to confirm the target cells before applying formatting or formulas. | Pass the address string to another API that requires a range reference in A1 notation. | Validate that a specific range exists before extracting or processing data.
// AI Prompts: Show C# code that loads an Excel workbook with Aspose.Cells, creates the range A1:C5, and returns its address. | Provide a robust Aspose.Cells example that handles a missing input file while still retrieving the address of a defined range. | Explain how to use the Range.Address property to get an A1‑style string for any cell block in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// A concise example that checks for an existing Excel file, creates a new workbook if needed, loads the first worksheet, defines the range A1:C5 using Aspose.Cells.Range, reads the Range.Address property, and prints the address string to the console.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists; if not, create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with a default worksheet
                workbook.Save(inputPath);   // optionally persist the empty workbook
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range that spans from A1 to C5 using Aspose.Cells.Range
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1", "C5");

            // Retrieve the address of the range
            string address = range.Address;

            // Output the address
            Console.WriteLine("Range address: " + address);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
