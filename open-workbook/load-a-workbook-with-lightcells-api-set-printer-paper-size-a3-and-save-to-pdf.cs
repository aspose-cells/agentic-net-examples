// Title: Load an Excel workbook (using LightCells API), set the first worksheet's printer paper size to A3, and save as PDF with Aspose.Cells for .NET
// AI Prompts: In C#, open an .xlsx file with Aspose.Cells LightCells API, change the first worksheet's PageSetup.PaperSize to PaperA3, and export the workbook to a PDF document. | Using Aspose.Cells for .NET, load a workbook, set the printer paper size of the first sheet to A3, and save the result as a PDF file.
// Common Searches: Aspose.Cells C# set worksheet paper size to A3 before PDF conversion | How to change page setup to A3 using LightCells API in Aspose.Cells | Export Excel to PDF with A3 layout using Aspose.Cells for .NET | C# Aspose.Cells set PaperSizeType.PaperA3 and save as PDF | Load large Excel file with LightCells and adjust page setup for PDF output
// Tags: Aspose.Cells LightCells load workbook | set worksheet paper size A3 | page setup PaperSizeType.PaperA3 C# | export workbook to PDF Aspose.Cells | adjust printer settings before PDF conversion | C# Aspose.Cells PDF export custom page layout

using System;
using Aspose.Cells;

// The example loads an Excel file, sets the first worksheet's printer paper size to A3 via PageSetup, and saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (standard loading; LightCells API can be used for large files,
        // but for setting page setup we need a Workbook object)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the printer paper size to A3 for the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
