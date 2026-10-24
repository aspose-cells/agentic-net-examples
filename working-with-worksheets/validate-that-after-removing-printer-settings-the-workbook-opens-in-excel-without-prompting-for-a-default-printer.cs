// Title: Clear printer settings from every worksheet in an Excel workbook using Aspose.Cells for .NET to stop default printer prompts
// Common Searches: Aspose.Cells C# remove printer settings from worksheets to avoid Excel printer dialog | How to disable default printer prompt when opening a workbook saved with Aspose.Cells | Set PageSetup.PrinterSettings to null for all sheets using Aspose.Cells .NET | Clear page setup printer configuration in an Excel file programmatically with Aspose.Cells | Prevent Excel from asking for a printer after saving workbook via Aspose.Cells
// Tags: clear worksheet printer settings Aspose.Cells | disable Excel printer dialog .NET | null PageSetup.PrinterSettings Aspose.Cells | remove printer configuration Excel workbook | Aspose.Cells save without printer prompt

using Aspose.Cells;
using System;

// The example loads an Excel file, iterates through each worksheet, sets the PageSetup.PrinterSettings property to null to clear any printer configuration, and saves the workbook. The resulting file opens in Excel without displaying a default printer selection dialog.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and clear printer settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Removing printer settings prevents Excel from prompting for a default printer
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
