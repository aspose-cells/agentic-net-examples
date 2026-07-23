// Title: Aspose.Cells .NET – Set Built‑in & Custom Document Properties and Export to PDF with Embedded Metadata
// Description: C# example that creates a workbook, assigns Author, Title, Subject and a custom property, updates the Author via the WorkbookMetadata API, and saves the workbook as a PDF where both built‑in and custom properties are written to the PDF Info dictionary using PdfSaveOptions.
// Keywords: Aspose.Cells set document properties C# | embed Excel metadata in PDF | WorkbookMetadata API Aspose.Cells | PdfSaveOptions custom properties | export custom document properties to PDF | Aspose.Cells PDF metadata example | C# Aspose.Cells document properties
// Common Searches: How to embed Excel document properties into a PDF with Aspose.Cells C#? | How to change the author of an Excel workbook using WorkbookMetadata? | How to export custom document properties to a PDF using Aspose.Cells? | Aspose.Cells PdfSaveOptions custom properties example | C# code to add built‑in and custom properties and save as PDF
// Developer Intent: Create an Excel file, modify its built‑in and custom metadata programmatically, and generate a PDF that carries those properties.
// Use Cases: Produce PDF reports that retain the original author, title, and project tags from the source spreadsheet. | Update workbook metadata after creation without opening the file in Excel. | Supply downstream document‑management systems with PDFs that contain searchable custom properties.
// AI Prompts: Generate C# code that adds several custom document properties to an Aspose.Cells workbook and exports them to the PDF Info dictionary. | Show how to read, modify, and save built‑in document properties of an existing Excel file using the WorkbookMetadata class. | Explain the effect of PdfCustomPropertiesExport.Standard on metadata embedded in the resulting PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Rendering;

// C# example that creates a workbook, assigns Author, Title, Subject and a custom property, updates the Author via the WorkbookMetadata API, and saves the workbook as a PDF where both built‑in and custom properties are written to the PDF Info dictionary using PdfSaveOptions.
class EmbedMetadataToPdf
{
    static void Main()
    {
        // Paths for intermediate Excel file and final PDF
        string excelPath = "DocumentWithMetadata.xlsx";
        string pdfPath = "DocumentWithMetadata.pdf";

        // -------------------------------------------------
        // 1. Create a new workbook and add some sample data
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // -------------------------------------------------
        // 2. Set built‑in document properties directly
        // -------------------------------------------------
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Demo Workbook";
        workbook.BuiltInDocumentProperties.Subject = "Metadata Embedding Example";

        // -------------------------------------------------
        // 3. Add a custom document property
        // -------------------------------------------------
        workbook.CustomDocumentProperties.Add("Project", "AsposeDemo");

        // -------------------------------------------------
        // 4. Save the workbook as an Excel file (required for metadata API)
        // -------------------------------------------------
        workbook.Save(excelPath);

        // -------------------------------------------------
        // 5. Load the workbook metadata, modify a built‑in property,
        //    and save the metadata back to the Excel file
        // -------------------------------------------------
        MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
        WorkbookMetadata metadata = new WorkbookMetadata(excelPath, metaOptions);

        // Example modification: change the Author property via metadata API
        metadata.BuiltInDocumentProperties.Author = "Jane Smith";

        // Save the updated metadata to the same Excel file
        metadata.Save(excelPath);

        // -------------------------------------------------
        // 6. Reload the workbook to ensure it contains the updated metadata
        // -------------------------------------------------
        Workbook updatedWorkbook = new Workbook(excelPath);

        // -------------------------------------------------
        // 7. Prepare PDF save options to export custom properties
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export custom document properties to the PDF's Info dictionary
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // -------------------------------------------------
        // 8. Save the workbook as PDF; built‑in and custom properties
        //    are embedded according to the options set above
        // -------------------------------------------------
        updatedWorkbook.Save(pdfPath, pdfOptions);

        Console.WriteLine("PDF generated with embedded metadata at: " + pdfPath);
    }
}
