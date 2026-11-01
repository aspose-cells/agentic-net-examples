// Title: Set a worksheet's print area and export it to PDF using Aspose.Cells for .NET
// Common Searches: how to define a custom print area for an Excel worksheet before PDF conversion with Aspose.Cells C# | Aspose.Cells .NET example for setting PageSetup.PrintArea and saving as PDF | C# code to limit printed range to A1:D10 when exporting workbook to PDF using Aspose.Cells | configure worksheet print settings in Aspose.Cells prior to PDF generation
// Tags: worksheet print area Aspose.Cells | PDF export from Excel Aspose.Cells | page setup print area C# | save workbook as PDF .NET | custom print range Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, fills cells A1 through D10 with data, assigns the range A1:D10 to the worksheet's PageSetup.PrintArea, and then saves the workbook directly as a PDF file using Aspose.Cells for .NET.
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
            sheet.Name = "Sheet1";

            // Populate some sample data (optional, just for illustration)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["D1"].PutValue("Header4");
            for (int row = 2; row <= 10; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"R{row - 1}C1");
                sheet.Cells[$"B{row}"].PutValue($"R{row - 1}C2");
                sheet.Cells[$"C{row}"].PutValue($"R{row - 1}C3");
                sheet.Cells[$"D{row}"].PutValue($"R{row - 1}C4");
            }

            // Define the print area (A1:D10)
            string printArea = "$A$1:$D$10";
            sheet.PageSetup.PrintArea = printArea;

            // Save the workbook as a PDF file
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
