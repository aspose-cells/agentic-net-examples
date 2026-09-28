// Title: Add custom document properties to an Excel workbook and export it as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, inserts custom document properties such as ProjectName, Version, Reviewed, CreatedBy, and ReviewDate, then saves the workbook as a PDF. | Write a C# snippet that sets several custom metadata fields on a workbook before calling Workbook.Save with SaveFormat.Pdf in Aspose.Cells.
// Common Searches: Aspose.Cells C# add custom document properties before converting Excel to PDF | How to embed custom metadata like ProjectName and ReviewDate in an Excel file and export to PDF with Aspose.Cells | C# example for setting workbook custom properties and saving as PDF using Aspose.Cells | Export Excel workbook to PDF while preserving custom document properties using Aspose.Cells .NET
// Tags: add custom properties Aspose.Cells | Excel to PDF conversion Aspose.Cells C# | set workbook metadata Aspose.Cells | SaveFormat.Pdf usage Aspose.Cells | C# workbook custom metadata

using System;
using Aspose.Cells;

// // Loads input.xlsx, adds custom document properties (ProjectName, Version, Reviewed, CreatedBy, ReviewDate), and saves the workbook as output.pdf in PDF format using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Add custom document properties (application‑specific data)
        // These properties can be retrieved later by other applications or scripts
        workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
        workbook.CustomDocumentProperties.Add("Version", "1.0.0");
        workbook.CustomDocumentProperties.Add("Reviewed", true);
        workbook.CustomDocumentProperties.Add("CreatedBy", "John Doe");
        workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);

        // Convert the workbook to PDF and save it
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
