// Title: Set a custom print area A1:M50 on a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, accesses the first worksheet, and assigns the PageSetup.PrintArea property to "A1:M50" before saving. | Show how to limit the printable range of an Excel sheet to cells A1 through M50 with the Aspose.Cells .NET API. | Generate a snippet that configures the print area for a worksheet and saves the file as an .xlsx using Aspose.Cells in C#.
// Common Searches: aspnet set print area A1:M50 using Aspose.Cells | how to define printable range in Aspose.Cells C# example | Aspose.Cells PageSetup.PrintArea property usage for specific cell range | C# code to limit Excel print area to A1:M50 with Aspose.Cells | example of setting print area in Aspose.Cells workbook
// Tags: Aspose.Cells PageSetup.PrintArea property | C# worksheet print area definition | Excel .xlsx printable range configuration | limit sheet printing Aspose.Cells | set printable area A1:M50 .NET

using Aspose.Cells;

// // This program creates a new workbook, accesses the first worksheet, sets the print area to cells A1 through M50 via PageSetup.PrintArea, and saves the file as PrintAreaExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the print area covering cells A1 through M50
        sheet.PageSetup.PrintArea = "A1:M50";

        // Save the workbook to a file
        workbook.Save("PrintAreaExample.xlsx");
    }
}
