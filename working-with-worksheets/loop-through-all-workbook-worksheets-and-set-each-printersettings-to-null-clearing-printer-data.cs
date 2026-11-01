// Title: How to clear printer settings for every worksheet in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX workbook with Aspose.Cells, iterates over all worksheets, and sets each sheet's PageSetup.PrinterSettings property to null. | Show how to save the workbook after removing printer configuration from every worksheet using Aspose.Cells. | Demonstrate how to clear printer settings only on worksheets whose names start with a specific prefix (e.g., "Report") with Aspose.Cells.
// Common Searches: aspnet remove printer settings from all worksheets in an Excel workbook | c# Aspose.Cells clear page setup printer configuration for each sheet | how to nullify PageSetup.PrinterSettings in a workbook using Aspose.Cells | reset printer data in Excel file programmatically with Aspose.Cells .NET
// Tags: clear worksheet printer settings Aspose.Cells | nullify PageSetup.PrinterSettings .NET | remove printer configuration from XLSX workbook | iterate worksheets Aspose.Cells clear printer data | reset printer settings Excel file Aspose.Cells

using Aspose.Cells;
using System;

// Loads an XLSX workbook, loops through each worksheet, sets PageSetup.PrinterSettings to null to erase printer information, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear printer settings by setting the PrinterSettings property to null
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the workbook after clearing printer data (replace with desired output path)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
