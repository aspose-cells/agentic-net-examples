// Title: Iterate over all named ranges in an Excel workbook and write each name with its RefersTo address to the debug log using Aspose.Cells for .NET
// Common Searches: how to enumerate named ranges and get their RefersTo addresses with Aspose.Cells in C# | Aspose.Cells C# example for writing named range names and addresses to Debug output | retrieve all defined names from an Excel file and log their cell references using Aspose.Cells | C# code to list workbook named ranges and output address to Visual Studio debug window | Aspose.Cells iterate workbook Names collection and print name and address
// Tags: named range address extraction Aspose.Cells | Aspose.Cells workbook defined names iteration | output named range details to Debug in C# | RefersTo property usage Aspose.Cells | log Excel defined names .NET

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample loads an Excel workbook, loops through the Worksheets.Names collection, reads each Name's RefersTo address, and writes the name and its address to the Visual Studio debug log.
class Program
{
    static void Main()
    {
        // Load the workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all named ranges defined in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // The RefersTo property contains the address the name points to (e.g., =Sheet1!$A$1:$B$2)
            string address = namedRange.RefersTo;

            // Output the name and its address to the debug log
            Debug.WriteLine($"Name: {namedRange.Text}, Address: {address}");
        }
    }
}
