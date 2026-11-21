// Title: Repeat a header row on every PDF page with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# repeat first worksheet row on each PDF page | How to set PrintTitleRows in Aspose.Cells before saving as PDF | Export Excel to PDF with repeating header rows using Aspose.Cells .NET | C# PageSetup.PrintTitleRows example for PDF output in Aspose.Cells
// Tags: Aspose.Cells PrintTitleRows PDF export | repeat header row PDF Aspose.Cells C# | PageSetup PrintTitleRows configuration Aspose.Cells | Excel to PDF with repeated headers using Aspose.Cells | C# Aspose.Cells PDF header repeat

using Aspose.Cells;

// The sample loads an Excel workbook, sets the first worksheet's PageSetup.PrintTitleRows to "$1:$1" so the header row repeats on each PDF page, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Set the first row as a print title so it repeats on every PDF page
        // "$1:$1" means rows 1 to 1 (the header row)
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Save the workbook as a PDF file with the repeated header rows
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
