// Title: Export an Excel workbook with updated charts to PDF while preserving chart formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook containing charts to PDF preserving formatting | How to keep Excel chart styles when saving as PDF with Aspose.Cells | Save Excel file with embedded charts as PDF using Aspose.Cells .NET API | C# code to convert .xlsx to .pdf while retaining chart appearance | Export updated chart workbook to PDF with Aspose.Cells SaveFormat.Pdf
// Tags: Aspose.Cells PDF conversion preserving chart styles | C# workbook to PDF with chart fidelity | SaveFormat.Pdf chart export example | Excel chart to PDF using Aspose.Cells | maintain chart appearance during PDF export

using Aspose.Cells;
using System;

// The example loads an existing workbook (input.xlsx) that includes an updated chart and saves it as a PDF (output.pdf) via Aspose.Cells, ensuring that all chart formatting and appearance are retained in the resulting PDF.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains the updated chart
        Workbook workbook = new Workbook("input.xlsx");

        // Export the workbook to PDF while preserving all chart formatting
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
