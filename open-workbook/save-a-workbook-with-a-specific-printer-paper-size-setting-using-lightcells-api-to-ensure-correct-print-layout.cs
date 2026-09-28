// Title: Set worksheet printer paper size to A4 and export as XLSX with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells PageSetup.PaperSize to assign PaperA4 to a worksheet, then save the workbook as an XLSX file in C#. | Modify the example to change the printer paper size to Letter and export the workbook to CSV using Aspose.Cells. | Add robust exception handling while configuring worksheet print layout and saving the file with a custom filename.
// Common Searches: C# code to apply A4 paper size to an Aspose.Cells worksheet prior to export | save Excel file with specific printer settings using Aspose.Cells .NET | how to change page setup paper size in Aspose.Cells workbook | C# example for configuring PageSetup.PaperSize and exporting to XLSX | Aspose.Cells set print layout for worksheet programmatically
// Tags: Aspose.Cells worksheet PageSetup configuration | C# set printer paper size A4 | save workbook as XLSX with custom print layout | Aspose.Cells export with specific paper size | configure Excel print settings programmatically

using System;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first worksheet, sets its PageSetup.PaperSize to A4, writes sample data to cells A1 and B1, and saves the workbook as Output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the printer paper size (e.g., A4) for the worksheet
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // (Optional) Add some sample data
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Save the workbook in XLSX format
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
