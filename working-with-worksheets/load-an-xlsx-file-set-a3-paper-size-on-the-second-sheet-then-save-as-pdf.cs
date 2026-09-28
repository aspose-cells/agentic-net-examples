// Title: How to set A3 paper size on the second worksheet of an XLSX file and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing XLSX workbook, changes the page setup of the second worksheet to A3 paper size, and saves the entire workbook as a PDF with Aspose.Cells. | Show a step‑by‑step example of adjusting the paper size of a specific worksheet before converting the workbook to PDF in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set A3 paper size on second sheet before PDF conversion | How to change page setup of a particular worksheet in Aspose.Cells | Convert XLSX to PDF with A3 paper size using Aspose.Cells .NET | Set custom paper size for a worksheet when saving workbook as PDF in C#
// Tags: worksheet page setup paper size Aspose.Cells | second worksheet A3 paper size C# | export workbook to PDF custom paper size Aspose.Cells | C# Aspose.Cells set paper size for specific sheet | save XLSX as PDF with A3 layout Aspose.Cells

using System;
using Aspose.Cells;

// Loads 'input.xlsx', sets the page setup of the second worksheet to A3 paper size, and saves the workbook as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // Set the paper size to A3 for the worksheet
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
