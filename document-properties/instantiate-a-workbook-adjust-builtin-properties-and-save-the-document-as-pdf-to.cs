using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // 1. Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose");
        sheet.Cells["A2"].PutValue("Metadata Demo");

        // 2. Save the workbook to a temporary Excel file
        string excelPath = "temp.xlsx";
        workbook.Save(excelPath);

        // 3. Prepare metadata options for document properties
        MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

        // 4. Load the workbook metadata using the provided constructor
        WorkbookMetadata metadata = new WorkbookMetadata(excelPath, metaOptions);

        // 5. Adjust built‑in document properties
        metadata.BuiltInDocumentProperties.Author = "John Doe";
        metadata.BuiltInDocumentProperties.Title = "Metadata Demo Workbook";
        metadata.BuiltInDocumentProperties.Subject = "Aspose.Cells Metadata Example";

        // 6. (Optional) Add a custom document property
        metadata.CustomDocumentProperties.Add("Category", "Demo");

        // 7. Save the modified metadata back to the same Excel file using the provided Save method
        metadata.Save(excelPath);

        // 8. Reload the workbook so it reflects the updated metadata
        Workbook updatedWorkbook = new Workbook(excelPath);

        // 9. Configure PDF save options to embed custom properties into the PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // 10. Save the workbook as PDF; built‑in and custom properties are now embedded
        updatedWorkbook.Save("output.pdf", pdfOptions);
    }
}