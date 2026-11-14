// Title: Delete all blank rows from an Excel worksheet and save the workbook as a PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove empty rows from an Excel file with Aspose.Cells before converting to PDF | How to compact an Excel worksheet by deleting blank rows using Aspose.Cells .NET API | Aspose.Cells DeleteBlankRows method example and PDF export in C# | Remove blank rows from a worksheet and generate PDF with Aspose.Cells for .NET | Delete blank rows in Excel and save as PDF using Aspose.Cells library
// Tags: Aspose.Cells DeleteBlankRows method | blank row deletion Aspose.Cells | worksheet PDF export Aspose.Cells | compact Excel workbook C# Aspose.Cells | save workbook as PDF Aspose.Cells

using Aspose.Cells;

// The sample loads 'input.xlsx' into an Aspose.Cells Workbook, accesses the first worksheet, removes all empty rows using the DeleteBlankRows() method, and then saves the resulting workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete all blank rows in the worksheet
        worksheet.Cells.DeleteBlankRows();

        // Save the compacted workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
