using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertWorkbookToPdfWithCustomProperties
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue("More Data");

        // Add custom document properties that will be exported to PDF
        workbook.CustomDocumentProperties.Add("Project", "Alpha");
        workbook.CustomDocumentProperties.Add("ReviewedBy", "Jane Doe");
        workbook.CustomDocumentProperties.Add("Revision", 3);
        workbook.CustomDocumentProperties.Add("Approved", true);

        // Create PDF save options and enable export of custom properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file with the custom properties included
        workbook.Save("WorkbookWithCustomProps.pdf", pdfOptions);
    }
}