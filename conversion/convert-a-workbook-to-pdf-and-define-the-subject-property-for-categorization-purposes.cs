// Title: Set the Subject built‑in document property and convert an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: how to assign a Subject built‑in document property to an Excel workbook before exporting to PDF with Aspose.Cells C# | Aspose.Cells C# convert XLSX to PDF while adding custom document properties | C# code example for setting Excel file Subject metadata and saving as PDF using Aspose.Cells | set built‑in document property Subject in Aspose.Cells and generate PDF | Aspose.Cells PDF export with custom Subject metadata in .NET
// Tags: Subject built‑in property API Aspose.Cells | Excel workbook PDF export with custom metadata C# | Aspose.Cells SaveFormat.Pdf with document properties | C# assign Excel Subject metadata prior to PDF export | Aspose.Cells document property handling for PDF categorization

using Aspose.Cells;
using System;

// The program loads 'input.xlsx', assigns 'Financial Report Q3' to the Subject built‑in document property, and saves the workbook as 'output.pdf' using the PDF save format.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Define the Subject property for categorization
        workbook.BuiltInDocumentProperties.Subject = "Financial Report Q3";

        // Convert and save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
