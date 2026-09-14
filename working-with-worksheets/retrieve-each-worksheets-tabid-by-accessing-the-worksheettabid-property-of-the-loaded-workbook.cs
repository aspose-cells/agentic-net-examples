// Title: Retrieve and print each worksheet's TabId from an Excel workbook using Aspose.Cells in C#
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, loops through all worksheets, and prints each sheet's name together with its TabId. | Show how to access the Worksheet.TabId property while enumerating workbook.Worksheets in Aspose.Cells for .NET. | Provide a console application example that loads a workbook, reads the TabId of every worksheet, and outputs the results.
// Common Searches: how to get worksheet TabId using Aspose.Cells C# | list all sheet identifiers (TabId) in an Excel file with Aspose.Cells | C# Aspose.Cells read TabId property for each worksheet | display worksheet names and TabId values in a console app using Aspose.Cells
// Tags: Aspose.Cells read worksheet TabId | C# iterate workbook worksheets Aspose.Cells | output worksheet TabId to console | load Excel file with Aspose.Cells | Worksheet.TabId property usage

using Aspose.Cells;
using System;

// Loads an Excel workbook with Aspose.Cells, iterates over each worksheet, reads the TabId property, and writes the worksheet name and its TabId to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the TabId of the current worksheet
            int tabId = sheet.TabId;

            // Output the worksheet name and its TabId
            Console.WriteLine($"Worksheet: {sheet.Name}, TabId: {tabId}");
        }
    }
}
