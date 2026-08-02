// Title: Get the address of range A1:C5 after loading an Excel workbook with Aspose.Cells for .NET (C#)
// Description: Loads or creates an Excel file, opens the first worksheet, creates a range for cells A1:C5, and prints the range’s address string to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# get range address | Excel range address .NET | Create range A1:C5 | Workbook load Aspose | Range.Address property | Aspose.Cells example
// Common Searches: Aspose.Cells get range address C# | How to read range address in Aspose.Cells | Create range A1:C5 Aspose.Cells .NET | Load workbook and retrieve range address | Range.Address property example
// Developer Intent: Load an Excel workbook and obtain the address string of the A1:C5 range using Aspose.Cells in C#.
// Use Cases: Display the range address in a console app for debugging purposes. | Verify that a specific range exists before applying formatting or formulas. | Log dynamically generated range addresses during batch processing of workbooks. | Include range addresses in generated reports or audit trails.
// AI Prompts: Show C# code to retrieve a range address using row and column indices instead of A1 notation. | Provide an example that gathers addresses of multiple ranges into a collection with Aspose.Cells. | Explain how to handle a missing workbook file without automatically creating a new one. | Demonstrate retrieving the address of a named range in an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRangeAddressDemo
{
    // Loads or creates an Excel file, opens the first worksheet, creates a range for cells A1:C5, and prints the range’s address string to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file to be loaded
            string inputFile = "input.xlsx";

            try
            {
                // Ensure the input file exists; create a new workbook if it does not
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"File \"{inputFile}\" not found. Creating a new workbook.");
                    Workbook tempWb = new Workbook();
                    tempWb.Save(inputFile);
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a Range object that represents cells A1:C5
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C5");

                // Display the address of the range
                Console.WriteLine("Range address: " + range.Address);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
