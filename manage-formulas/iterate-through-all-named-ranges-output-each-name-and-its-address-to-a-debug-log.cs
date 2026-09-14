// Title: Log each named range and its address from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an .xlsx file with Aspose.Cells, iterates the workbook's NameCollection, and writes each range's Text and RefersTo to the Debug console. | Generate a .NET method that returns a collection of (rangeName, rangeAddress) tuples by scanning all defined names in a workbook via Aspose.Cells. | Create a snippet that uses Aspose.Cells to list all named ranges in a spreadsheet and outputs their addresses using System.Diagnostics.Debug.
// Common Searches: how to enumerate defined names in an Excel file with Aspose.Cells C# | Aspose.Cells get address of each named range programmatically | C# debug write named range references from workbook using Aspose.Cells | list all workbook names and references Aspose.Cells .NET example | retrieve named range Text and RefersTo Aspose.Cells C#
// Tags: Aspose.Cells enumerate named ranges | C# debug output Aspose.Cells | retrieve named range address .NET | list workbook defined names Aspose.Cells | iterate NameCollection Excel

using System;
using System.Diagnostics;
using Aspose.Cells;

// Loads 'input.xlsx' with Aspose.Cells, accesses workbook.Worksheets.Names, loops through each Name, and writes its Text and RefersTo values to the Debug log.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of all named ranges in the workbook
        NameCollection namedRanges = workbook.Worksheets.Names;

        // Iterate through each named range
        foreach (Name namedRange in namedRanges)
        {
            // Retrieve the name of the range
            string rangeName = namedRange.Text;

            // Retrieve the address (reference) of the range
            string rangeAddress = namedRange.RefersTo;

            // Output the name and address to the debug log
            Debug.WriteLine($"Name: {rangeName}, Address: {rangeAddress}");
        }
    }
}
