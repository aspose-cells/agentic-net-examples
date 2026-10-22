// Title: Retrieve the TabId of each worksheet in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells and prints every worksheet's name together with its TabId. | Show how to loop through a Workbook's Worksheets collection and access the Worksheet.TabId property for each sheet. | Create a C# example that stores worksheet names and their TabIds in a Dictionary while using Aspose.Cells.
// Common Searches: asp.net c# get worksheet tabid using aspose.cells | how to list tab identifiers for all sheets in an Excel file with Aspose.Cells | c# retrieve tabid of each worksheet in a workbook asp.net core | asp.net read worksheet TabId property from xlsx file | example code to display sheet names and tabids with Aspose.Cells
// Tags: read worksheet TabId Aspose.Cells | iterate workbook worksheets C# | Aspose.Cells get sheet identifier | C# extract Excel sheet TabId | Aspose.Cells TabId property usage

using Aspose.Cells;
using System;

// The program loads an Excel file with Aspose.Cells, iterates through all worksheets, reads each sheet's TabId property, and writes the worksheet name and TabId to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and retrieve its TabId
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;
            // Output the worksheet name and its TabId
            Console.WriteLine($"Worksheet: {sheet.Name}, TabId: {tabId}");
        }
    }
}
