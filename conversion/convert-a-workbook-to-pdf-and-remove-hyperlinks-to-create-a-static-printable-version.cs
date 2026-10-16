// Title: Convert an Excel workbook to a static printable PDF by removing all hyperlinks with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove all hyperlinks before exporting to PDF | How to generate a PDF from Excel without clickable links using Aspose.Cells | C# code to clear worksheet hyperlinks and save as static PDF with Aspose.Cells | Export .xlsx to PDF with hyperlinks disabled in Aspose.Cells .NET | Create a printable PDF from Excel by stripping hyperlinks in C#
// Tags: clear worksheet hyperlinks Aspose.Cells | export workbook to PDF Aspose.Cells C# | hyperlink removal before PDF conversion | static printable PDF generation Aspose.Cells | SaveFormat.Pdf usage Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel file, iterates through each worksheet to clear its Hyperlinks collection, and then saves the workbook as a PDF using Aspose.Cells, resulting in a static, printable PDF with all hyperlinks removed.
class WorkbookToPdfStatic
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and remove all hyperlinks
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear the Hyperlinks collection to make the sheet static
            sheet.Hyperlinks.Clear();
        }

        // Save the modified workbook as a PDF (static, printable version)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
