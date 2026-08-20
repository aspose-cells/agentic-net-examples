// Title: C# – Load XLS, Evaluate Formulas, and Save Encrypted PDF (Owner Password) with Aspose.Cells
// Description: Load an XLS workbook using Aspose.Cells, trigger formula calculation, apply an owner‑only password via PdfSecurityOptions, and export the result as a protected PDF file.
// Keywords: Aspose.Cells XLS to PDF | calculate formulas before PDF export | PDF owner password encryption | PdfSaveOptions CalculateFormula | PdfSecurityOptions C# | secure PDF generation from Excel
// Common Searches: Aspose.Cells export Excel to password protected PDF C# | how to calculate formulas when saving workbook as PDF | set owner password for PDF using Aspose.Cells | encrypt PDF generated from XLS with Aspose.Cells
// Developer Intent: Create a PDF from an XLS file, ensure all Excel formulas are evaluated, and protect the PDF with an owner password.
// Use Cases: Generate confidential PDF reports that reflect Excel calculations. | Batch‑convert multiple XLS workbooks into encrypted PDFs for secure distribution. | Produce invoice PDFs from Excel templates while preventing unauthorized editing.
// AI Prompts: Write C# code with Aspose.Cells to load an .xls file, evaluate all formulas, and save it as a PDF protected by an owner password. | Show how to configure additional PDF permissions (printing, copying) when encrypting a PDF using Aspose.Cells PdfSecurityOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Load an XLS workbook using Aspose.Cells, trigger formula calculation, apply an owner‑only password via PdfSecurityOptions, and export the result as a protected PDF file.
class Program
{
    static void Main()
    {
        // Path to the source XLS file
        string sourceFile = "input.xls";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourceFile);

        // Create PDF save options and enable formula calculation
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.CalculateFormula = true; // calculate formulas before saving

        // Set PDF security options with an owner password
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.OwnerPassword = "ownerPassword123"; // owner password
        // (optional) you can also set a user password or permissions here
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        string outputPdf = "encrypted_output.pdf";
        workbook.Save(outputPdf, pdfSaveOptions);
    }
}
