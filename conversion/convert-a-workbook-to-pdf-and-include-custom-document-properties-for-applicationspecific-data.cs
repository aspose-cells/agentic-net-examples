// Title: Convert an Excel workbook to PDF and embed custom document properties using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add custom document properties before saving workbook as PDF | How to include custom metadata in PDF generated from Excel with Aspose.Cells | Aspose.Cells PDF export with custom properties like author and report ID in C# | Embedding application‑specific data into PDF metadata during Excel to PDF conversion using Aspose.Cells | Save Excel workbook as PDF with custom fields using Aspose.Cells .NET API
// Tags: add custom document properties Aspose.Cells PDF | save workbook as PDF with metadata C# | Aspose.Cells PDF conversion custom properties | embed application specific data PDF Aspose.Cells | C# Excel to PDF custom metadata

using System;
using Aspose.Cells;

// The sample creates a Workbook, fills it with sample data, adds custom document properties (ReportAuthor, ReportId, IsConfidential, GeneratedOn), and saves the workbook as a PDF where those properties are embedded in the PDF metadata.
class WorkbookToPdfWithCustomProperties
{
    static void Main()
    {
        // Load an existing workbook or create a new one
        // Here we create a new workbook for demonstration
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Oranges");
        sheet.Cells["B3"].PutValue(85);

        // Add custom document properties (application‑specific data)
        // Example: add a property for the report author and a custom ID
        workbook.CustomDocumentProperties.Add("ReportAuthor", "John Doe");
        workbook.CustomDocumentProperties.Add("ReportId", 20231001);
        workbook.CustomDocumentProperties.Add("IsConfidential", true);
        workbook.CustomDocumentProperties.Add("GeneratedOn", DateTime.Now);

        // Save the workbook as PDF; custom properties are embedded in the PDF metadata
        workbook.Save("ReportWithProperties.pdf", SaveFormat.Pdf);

        Console.WriteLine("Workbook converted to PDF with custom document properties.");
    }
}
