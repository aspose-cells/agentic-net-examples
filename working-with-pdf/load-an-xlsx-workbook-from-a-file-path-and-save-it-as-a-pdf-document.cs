// Title: Load an XLSX workbook from a file path and save it as PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load Excel file from disk and export to PDF | Convert local XLSX to PDF using Aspose.Cells .NET API | Save workbook as PDF in C# with Aspose.Cells SaveFormat.Pdf | Example code for converting XLSX to PDF with Aspose.Cells in a console app | How to use Aspose.Cells to generate PDF from an existing Excel file in C#
// Tags: Aspose.Cells load workbook from file path | Aspose.Cells save workbook as PDF | C# Excel to PDF conversion using Aspose.Cells | SaveFormat.Pdf parameter Aspose.Cells | XLSX to PDF export .NET

using Aspose.Cells;

// Creates a Workbook object from a local XLSX file and saves it directly as a PDF using Aspose.Cells' SaveFormat.Pdf option.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from the specified file path
        Workbook workbook = new Workbook("input.xlsx");

        // Save the loaded workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
