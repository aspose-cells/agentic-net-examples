// Title: Insert a new column at index 3 in an Excel worksheet and export the sheet to PDF using Aspose.Cells for .NET
// AI Prompts: Insert a column at zero‑based index 3 in the first worksheet of an Excel file and save the workbook as a PDF with Aspose.Cells. | Add several columns before column D, then generate a PDF from the updated workbook using Aspose.Cells in C#.
// Common Searches: asp.net insert column at index 3 using Aspose.Cells before PDF export | c# how to add a column to an existing Excel file and convert it to PDF with Aspose.Cells | insert column D in worksheet and save workbook as PDF using Aspose.Cells library
// Tags: insert column Aspose.Cells C# | export worksheet to PDF Aspose.Cells | add column before PDF conversion .NET | modify Excel file structure programmatically | save workbook as PDF after column insertion

using System;
using Aspose.Cells;

// The example loads 'input.xlsx', accesses the first worksheet, inserts a new column at zero‑based index 3 (before column D), and saves the workbook as 'output.pdf' in PDF format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a new column at index 3 (zero‑based, i.e., before column D)
        worksheet.Cells.InsertColumn(3);

        // Export the entire workbook (or specific worksheet) as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
