// Title: Convert an Excel workbook to PDF with French language settings using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with French locale | How to set workbook culture to fr-FR before saving as PDF in .NET | C# convert .xlsx to PDF and specify document language using Aspose.Cells | Set PDF output language to French when converting Excel with Aspose.Cells | Aspose.Cells PDF conversion multilingual support C# example
// Tags: CultureInfo based PDF export in Aspose.Cells | Workbook.Settings.CultureInfo French locale | C# multilingual workbook PDF generation | Aspose.Cells document language configuration | SaveFormat.Pdf locale handling

using System;
using System.Globalization;
using Aspose.Cells;

// The code loads an Excel file, sets the workbook's CultureInfo to French (fr-FR) via Settings.CultureInfo, and saves the workbook as a PDF using Aspose.Cells.
class WorkbookToPdf
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set the document language to French (France)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Convert and save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
