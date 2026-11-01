// Title: Set A3 paper size on the second worksheet of an XLSX file and export it to PDF with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX workbook, change the PaperSize of the second worksheet to A3, and save the result as a PDF using Aspose.Cells in C#. | Adjust the page setup of a specific worksheet to A3 and then convert the workbook to PDF format with Aspose.Cells for .NET. | Use Aspose.Cells to apply an A3 paper size to worksheet index 1 before generating a PDF file.
// Common Searches: Aspose.Cells C# set paper size for worksheet index 1 | Convert second sheet of Excel to PDF with A3 page size using Aspose.Cells | How to change page setup of a specific worksheet before PDF export in Aspose.Cells .NET | C# Aspose.Cells export selected worksheet to PDF with custom paper size | Aspose.Cells set A3 page size on worksheet and save as PDF
// Tags: set worksheet paper size Aspose.Cells | export specific worksheet to PDF C# | A3 page setup Aspose.Cells | convert XLSX to PDF custom page size | modify page setup before PDF conversion Aspose.Cells

using System;
using Aspose.Cells;

// The example loads 'input.xlsx', sets the PaperSize of the second worksheet (index 1) to A3, and saves the workbook as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1) and set its paper size to A3
        Worksheet sheet = workbook.Worksheets[1];
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
