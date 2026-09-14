// Title: Load an Excel workbook and retrieve the address string of range A1:C5 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, creates a range covering cells A1 through C5 on the first worksheet, and prints its address as a string. | Show how to use CellsHelper in Aspose.Cells to convert a range's first row/column and dimensions into the standard A1:C5 address format. | Provide a robust example that checks for file existence, loads the workbook, and extracts the address of any specified range using Aspose.Cells.
// Common Searches: asp.net aspose.cells get address of cells A1 through C5 in C# | how to convert Aspose.Cells range indices to Excel address string C# | extract address of a selected range from worksheet using Aspose.Cells | C# example for retrieving range address with Aspose.Cells
// Tags: Aspose.Cells range address extraction | C# Aspose.Cells define custom cell range | Aspose.Cells CellsHelper address conversion | load Excel workbook using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, creates a range covering A1:C5, uses CellsHelper to translate the range's first row/column and size into start and end cell names, builds the address string "A1:C5", and prints it to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers cells A1 to C5
            AsposeRange range = sheet.Cells.CreateRange("A1:C5");

            // Build the address string of the range using its properties
            string startCell = CellsHelper.CellIndexToName(range.FirstRow, range.FirstColumn);
            string endCell = CellsHelper.CellIndexToName(
                range.FirstRow + range.RowCount - 1,
                range.FirstColumn + range.ColumnCount - 1);
            string address = $"{startCell}:{endCell}";

            // Display the address
            Console.WriteLine(address);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
