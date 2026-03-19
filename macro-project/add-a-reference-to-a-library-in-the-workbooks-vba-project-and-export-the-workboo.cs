using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceAndExportPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure the workbook has a VBA project (required for adding references)
        if (workbook.VbaProject == null)
        {
            // Save as a macro‑enabled workbook to create the VBA project
            string tempXlsm = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            workbook.Save(tempXlsm, SaveFormat.Xlsm);

            // Reload the workbook which now contains a VBA project
            workbook = new Workbook(tempXlsm);

            // Clean up the temporary file
            File.Delete(tempXlsm);
        }

        // Add a reference to an Automation type library (e.g., stdole)
        VbaProject vbaProject = workbook.VbaProject;
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation"
        );

        // Prepare PDF save options (example: export document structure)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Export the workbook to PDF
        workbook.Save("WorkbookWithVbaReference.pdf", pdfOptions);

        Console.WriteLine("PDF exported successfully.");
    }
}