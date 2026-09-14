// Title: Assign A4 paper size to data worksheets and Letter size to summary worksheets using Aspose.Cells for .NET
// AI Prompts: Set the PageSetup.PaperSize to A4 for every worksheet whose name includes "Data" and to Letter for worksheets whose name includes "Summary" in a C# Aspose.Cells workbook. | Write C# code that loops through all sheets in an Excel file and applies different paper sizes based on each sheet's role with Aspose.Cells. | Programmatically change the page setup of an existing workbook so that data sheets use A4 and summary sheets use Letter paper size using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells set A4 paper size for sheets containing Data | How to apply Letter paper size to summary worksheets with Aspose.Cells C# | Change page setup per worksheet based on name using Aspose.Cells .NET | C# example for conditional paper size assignment in Excel with Aspose.Cells
// Tags: set worksheet paper size Aspose.Cells | conditional page setup by sheet name | A4 paper size for data sheets Aspose.Cells | Letter paper size for summary sheets Aspose.Cells | iterate worksheets C# Aspose.Cells | apply different page sizes per sheet

using System;
using Aspose.Cells;

// Loads 'input.xlsx', iterates through each worksheet, assigns A4 paper size to sheets whose name contains "Data" and Letter size to sheets whose name contains "Summary", then saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load existing workbook (using provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If worksheet name indicates a data sheet, set A4 size
            if (sheet.Name.IndexOf("Data", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }
            // If worksheet name indicates a summary sheet, set Letter size
            else if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }
        }

        // Save the modified workbook (using provided save rule)
        workbook.Save("output.xlsx");
    }
}
