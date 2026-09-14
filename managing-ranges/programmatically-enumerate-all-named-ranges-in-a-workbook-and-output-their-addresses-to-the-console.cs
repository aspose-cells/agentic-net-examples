// Title: List all named ranges in an Excel workbook and print their addresses with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loads an .xlsx file and prints each defined name together with its RefersTo address to the console. | Generate a console application that iterates over workbook.Worksheets.Names and displays the name, worksheet, and range address for every named range. | Create a C# snippet that enumerates all named ranges in a workbook and outputs the name, sheet reference, and address in a formatted line.
// Common Searches: Aspose.Cells C# list all defined names in an Excel file | how to retrieve RefersTo address of named ranges using Aspose.Cells .NET | C# console output of Excel named range addresses with Aspose.Cells | enumerate workbook named ranges programmatically Aspose.Cells
// Tags: Aspose.Cells enumerate named ranges | C# list defined names RefersTo | output named range addresses console | workbook.Worksheets.Names iteration | retrieve Excel named range addresses .NET

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, iterates through workbook.Worksheets.Names, and writes each named range's name and RefersTo address to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all defined names (named ranges) in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // The RefersTo property contains the address of the range (e.g., =Sheet1!$A$1:$B$2)
            string address = namedRange.RefersTo;

            // Output the name and its address to the console
            Console.WriteLine($"{namedRange.Text}: {address}");
        }
    }
}
