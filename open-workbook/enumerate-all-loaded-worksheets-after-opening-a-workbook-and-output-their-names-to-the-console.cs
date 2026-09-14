// Title: List all worksheet names from an Excel file using Aspose.Cells for .NET (C#) and print to console
// AI Prompts: Write a C# console program that opens a .xlsx workbook with Aspose.Cells and prints each worksheet's Name property. | Create a reusable C# method that takes a file path, loads the workbook via Aspose.Cells, and returns a list of worksheet titles. | Show how to loop through the Workbook.Worksheets collection in Aspose.Cells for .NET and output each sheet name to standard output.
// Common Searches: asp.net console app list worksheet names from Excel using Aspose.Cells | c# code to read all sheet titles from an .xlsx file with Aspose.Cells library | how to enumerate worksheets in a workbook and display their names in C#
// Tags: Aspose.Cells enumerate workbook worksheets | C# console output Excel sheet names | Aspose.Cells open Excel file | Aspose.Cells worksheet collection traversal | retrieve worksheet Name property

using System;
using Aspose.Cells;

// // Loads 'input.xlsx' with Aspose.Cells, iterates over workbook.Worksheets, and writes each worksheet's Name to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Output the name of each worksheet to the console
            Console.WriteLine(sheet.Name);
        }
    }
}
