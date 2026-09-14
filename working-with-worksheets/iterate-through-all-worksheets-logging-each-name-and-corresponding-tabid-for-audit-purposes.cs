// Title: Log each worksheet’s name and TabId while iterating through all worksheets in an Aspose.Cells workbook (C#)
// AI Prompts: Write a C# routine using Aspose.Cells that prints the Name and TabId of every worksheet in a given workbook. | Create a .NET method that returns a list of (worksheet name, TabId) pairs from an Aspose.Cells workbook. | Generate C# code that iterates over a workbook’s worksheets and writes each sheet’s name and TabId to a log file with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to get worksheet TabId while looping through worksheets | C# iterate all sheets in Excel file and output their TabId using Aspose.Cells | retrieve worksheet identifiers for audit in Aspose.Cells .NET | log worksheet names with TabId for Excel workbook in C# Aspose.Cells
// Tags: Aspose.Cells iterate worksheets | Aspose.Cells get worksheet TabId | C# log worksheet identifiers | Excel workbook audit Aspose.Cells | Aspose.Cells worksheet enumeration logging

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, loops through each worksheet, and writes the worksheet’s Name and TabId to the console for auditing purposes.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Log the worksheet name and its TabId for audit purposes
            Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
        }
    }
}
