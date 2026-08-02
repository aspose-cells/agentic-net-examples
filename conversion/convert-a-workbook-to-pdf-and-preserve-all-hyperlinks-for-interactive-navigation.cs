// Title: C# – Convert Excel to PDF with Clickable Hyperlinks using Aspose.Cells
// Description: Learn how to use Aspose.Cells for .NET to save an Excel workbook as a PDF while preserving all hyperlinks. The example shows creating a PdfSaveOptions object, enabling ExportDocumentStructure, and generating a PDF with fully functional links.
// Keywords: Aspose.Cells PDF conversion | C# preserve Excel hyperlinks | PdfSaveOptions ExportDocumentStructure | Excel to PDF clickable links | Aspose.Cells hyperlink support | Convert .xlsx to PDF .NET | ExportDocumentStructure true
// Common Searches: Aspose.Cells keep hyperlinks when saving as PDF | ExportDocumentStructure property PDF hyperlink | C# convert Excel to PDF with active links | How to retain Excel hyperlinks in PDF using Aspose.Cells | PdfSaveOptions hyperlink preservation example
// Developer Intent: Generate a PDF from an Excel workbook and ensure every hyperlink remains clickable in the output.
// Use Cases: Create PDF reports from spreadsheets where users can navigate to external sites or internal sections via active links. | Produce downloadable manuals or catalogs from Excel data that keep embedded hyperlink navigation intact. | Automate batch conversion of multiple .xlsx files to PDFs for a web portal while preserving link functionality.
// AI Prompts: Write C# code using Aspose.Cells to save an Excel workbook as a PDF with ExportDocumentStructure enabled so that all hyperlinks stay clickable. | Explain the role of PdfSaveOptions.ExportDocumentStructure in preserving hyperlinks during Excel‑to‑PDF conversion. | Show how to combine ExportDocumentStructure with other PDF settings such as page orientation, image quality, and password protection in Aspose.Cells.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells for .NET to save an Excel workbook as a PDF while preserving all hyperlinks. The example shows creating a PdfSaveOptions object, enabling ExportDocumentStructure, and generating a PDF with fully functional links.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Preserve hyperlinks for interactive navigation in the PDF
        // ExportDocumentStructure includes link information such as bookmarks and hyperlinks
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
