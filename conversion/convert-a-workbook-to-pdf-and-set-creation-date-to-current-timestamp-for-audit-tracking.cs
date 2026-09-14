// Title: Convert an Excel workbook to PDF and set the PDF creation date to the current time with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads or creates a Workbook, sets PdfSaveOptions.CreatedTime to DateTime.Now, and saves the workbook as a PDF using Aspose.Cells. | Show how to embed a dynamic creation timestamp into a PDF generated from an Excel file with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells how to set PDF CreatedTime property when saving workbook | C# save Excel as PDF with current timestamp for audit tracking | PdfSaveOptions CreatedTime example in Aspose.Cells .NET | Add creation date metadata to PDF generated from Excel using Aspose
// Tags: Aspose.Cells PdfSaveOptions CreatedTime | Excel to PDF conversion audit metadata | C# set PDF creation timestamp Aspose | Workbook save as PDF with dynamic date

using System;
using Aspose.Cells;

// // Creates a Workbook, optionally writes data, configures PdfSaveOptions.CreatedTime with DateTime.Now, and saves the workbook as 'AuditDocument.pdf' to embed the current creation timestamp for audit purposes.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Optional: add some data to the workbook
        workbook.Worksheets[0].Cells["A1"].PutValue("Audit PDF Example");

        // Create PDF save options and set the creation timestamp
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.Now // current time for audit tracking
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("AuditDocument.pdf", pdfOptions);
    }
}
