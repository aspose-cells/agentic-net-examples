// Title: Remove a worksheet's embedded printer configuration in Aspose.Cells (C#) by setting PageSetup.PrinterSettings to null
// AI Prompts: Generate C# code that creates a workbook, accesses the first worksheet, clears its printer configuration by assigning null to PageSetup.PrinterSettings, and saves the file. | Demonstrate how to nullify the PageSetup.PrinterSettings property to prevent printer data from being stored in an Aspose.Cells workbook. | Provide a snippet that nullifies Worksheet.PageSetup.PrinterSettings in Aspose.Cells to remove printer metadata before saving.
// Common Searches: Aspose.Cells C# how to clear printer settings on a worksheet before saving | set worksheet PageSetup.PrinterSettings to null Aspose.Cells example | remove embedded printer configuration from Excel workbook using Aspose.Cells API | C# code to delete printer configuration from Aspose.Cells worksheet
// Tags: reset worksheet printer settings Aspose.Cells C# | disable PageSetup.PrinterSettings Aspose.Cells | clear printer metadata workbook Aspose.Cells | disable printer settings in Excel workbook Aspose.Cells | Aspose.Cells worksheet printer setup cleanup

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PrinterSettings property to null to clear any embedded printer configuration, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Remove the embedded printer configuration by setting PrinterSettings to null
            sheet.PageSetup.PrinterSettings = null;

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
