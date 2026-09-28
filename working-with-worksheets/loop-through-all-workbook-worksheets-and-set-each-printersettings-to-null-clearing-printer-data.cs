// Title: How to clear printer settings for every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, loops through all worksheets, and sets each sheet's PageSetup.PrinterSettings property to null. | Generate a .NET example that removes printer configuration from every worksheet in a workbook and saves the modified file under a new name. | Provide a step‑by‑step Aspose.Cells snippet to reset printer settings across all sheets in an existing Excel workbook.
// Common Searches: Aspose.Cells C# clear printer settings for all worksheets in a workbook | remove page setup printer configuration from each sheet using Aspose.Cells .NET | how to reset printer settings in every worksheet of an Excel file with Aspose.Cells
// Tags: Aspose.Cells remove worksheet printer configuration | set PageSetup.PrinterSettings to null | loop through worksheets to reset printer settings | save workbook after clearing printer data | C# Aspose.Cells printer settings cleanup

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', iterates over each worksheet, assigns null to sheet.PageSetup.PrinterSettings to clear printer data, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear the printer settings for the current worksheet
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
