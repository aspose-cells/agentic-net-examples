// Title: How to convert an Excel workbook to PDF and set the Author built‑in document property using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, assigns a value to the Author built‑in document property, adds sample cells, and saves the workbook as a PDF with Aspose.Cells. | Show how to embed author metadata into an Excel file before exporting it to PDF using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells set workbook author property before PDF export | How to add custom document properties to Excel file and convert to PDF with Aspose.Cells | Saving Excel as PDF with author metadata using Aspose.Cells .NET | Aspose.Cells built‑in document properties example for PDF conversion
// Tags: Aspose.Cells assign author built‑in property | convert workbook to PDF using Aspose.Cells C# | embed author metadata in Excel PDF export | C# document properties for PDF conversion Aspose.Cells | save Excel as PDF with custom author metadata

using System;
using Aspose.Cells;

// // This example creates a new Workbook, sets the Author built‑in document property to "John Doe", writes sample values to cells, and saves the workbook as a PDF file using Aspose.Cells for .NET.
class ConvertWorkbookToPdfWithAuthor
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the author property using built‑in document properties
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

        // Add some sample data to demonstrate the workbook content
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Workbook");
        sheet.Cells["A2"].PutValue("Created by: " + workbook.BuiltInDocumentProperties["Author"].Value);
        sheet.Cells["A3"].PutValue(DateTime.Now.ToString());

        // Convert the workbook to PDF and save it to disk
        workbook.Save("SampleOutput.pdf", SaveFormat.Pdf);
    }
}
