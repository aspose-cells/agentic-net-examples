// Title: Assign a custom integer TabId to a specific worksheet and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set the TabId property of a chosen worksheet to a specific integer and then write the workbook to a new file. | Programmatically change the tab identifier of an existing Excel sheet with Aspose.Cells and persist the changes by saving the workbook. | Update the TabId of a worksheet named 'Sheet1' in an input.xlsx file and export the modified workbook as output.xlsx using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to change worksheet TabId value | set Excel sheet tab identifier programmatically with Aspose.Cells | example of assigning TabId to a worksheet and saving workbook in .NET | modify worksheet TabId property and export new Excel file using Aspose.Cells
// Tags: worksheet TabId assignment Aspose.Cells | update Excel sheet TabId C# | save workbook after TabId change Aspose.Cells | modify worksheet tab identifier programmatically | Aspose.Cells set TabId property example

using System;
using Aspose.Cells;

// The code loads an existing Excel file, sets the TabId of the specified worksheet to a custom integer, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the target worksheet (by name or index)
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name or use workbook.Worksheets[0]

        // Assign a new integer TabId to the worksheet
        worksheet.TabId = 12345; // set desired TabId value

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
