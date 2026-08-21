// Title: Hide First Worksheet Tab and Save as New Excel File with Aspose.Cells for .NET (C#)
// Description: Load an existing workbook, set the first worksheet's IsVisible property to false to hide its tab, and save the modified workbook as a new XLSX file using Aspose.Cells in C#.
// Keywords: Aspose.Cells hide worksheet tab | C# hide first sheet | set worksheet visibility Aspose | export workbook with hidden sheet | save Excel file invisible tab
// Common Searches: Aspose.Cells hide first worksheet tab C# | how to make a worksheet tab invisible with Aspose.Cells | save Excel workbook with hidden sheet using .NET | C# hide worksheet tab and export file | Aspose.Cells hide sheet before saving
// Developer Intent: Hide the first worksheet tab and export the workbook to a new Excel file using Aspose.Cells in C#.
// Use Cases: Distribute a report while keeping the index or configuration sheet hidden from end users. | Create a template that contains a hidden data sheet for macros, then generate a clean version for clients. | Publish a workbook on a public portal with internal calculation sheets concealed for security.
// AI Prompts: Generate C# code with Aspose.Cells that hides the first worksheet tab and saves the workbook as output.xlsx. | Explain how to hide multiple worksheets, preserve their visibility states, and export the workbook using Aspose.Cells. | Provide robust error handling for missing input files and permission issues when hiding a worksheet and saving the workbook.

using System;
using Aspose.Cells;

// Load an existing workbook, set the first worksheet's IsVisible property to false to hide its tab, and save the modified workbook as a new XLSX file using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Hide the first worksheet tab (index 0)
        // You can also use SetVisible(false, true) or VisibilityType = VisibilityType.Hidden
        workbook.Worksheets[0].IsVisible = false;

        // Save the workbook to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
