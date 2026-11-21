// Title: Enumerate and print all worksheet names from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to list all sheet names in an Excel file using Aspose.Cells | How to iterate through Workbook.Worksheets and get names in Aspose.Cells | Print worksheet names to console after loading workbook with Aspose.Cells .NET | Retrieve Excel worksheet collection names in a .NET application using Aspose.Cells | Sample program to display all worksheet names from input.xlsx with Aspose.Cells
// Tags: enumerate worksheet names Aspose.Cells .NET | Workbook.Worksheets iteration C# | output Excel sheet names to console | load workbook and list sheets Aspose.Cells | C# Aspose.Cells worksheet collection enumeration

using System;
using Aspose.Cells;

// The example loads an Excel file (input.xlsx) into an Aspose.Cells Workbook, iterates over the Workbook.Worksheets collection, and writes each Worksheet.Name to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enumerate all worksheets that are loaded in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Output each worksheet's name to the console
            Console.WriteLine(sheet.Name);
        }
    }
}
