// Title: Delete all blank rows from an Excel worksheet and save the workbook as a PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, removes every completely empty row from the first worksheet, and exports the workbook to a PDF file. | Show how to invoke Worksheet.Cells.DeleteBlankRows and then save the updated workbook as a PDF in a .NET console program.
// Common Searches: aspnet delete blank rows from Excel worksheet using Aspose.Cells | how to remove empty rows before converting Excel to PDF with Aspose.Cells C# | Aspose.Cells DeleteBlankRows method example for PDF export | C# clean Excel data and generate PDF using Aspose.Cells | remove rows with no data in Aspose.Cells and save as PDF
// Tags: blank row removal using Aspose.Cells C# | Aspose.Cells PDF generation after row removal | remove empty rows Excel Aspose.Cells | generate PDF from cleaned worksheet Aspose.Cells | compact worksheet before PDF conversion Aspose.Cells

using Aspose.Cells;
using System;

// Loads 'input.xlsx', deletes all rows that contain no data in any cell from the first worksheet using DeleteBlankRows, and saves the resulting workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify by name/index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete all blank rows in the worksheet
        // This method removes rows that contain no data in any cell
        worksheet.Cells.DeleteBlankRows();

        // Save the compacted workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
