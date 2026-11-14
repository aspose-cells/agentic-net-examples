// Title: Delete rows 30‑42, insert five new rows, and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# delete rows 30 to 42 and insert new rows | C# code to remove a range of rows and export Excel to PDF with Aspose.Cells | Insert rows with default height after deleting rows using Aspose.Cells .NET | How to delete multiple rows and then save workbook as PDF in Aspose.Cells | Aspose.Cells example for row range deletion and PDF conversion in C#
// Tags: delete rows range Aspose.Cells C# | insert rows default height Aspose.Cells | export workbook to PDF Aspose.Cells | row manipulation Excel Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample loads an Excel file, removes rows 30‑42 from the first worksheet, inserts five new rows at that position with the default row height, and then saves the modified workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 30 to 42 (1‑based). Aspose.Cells uses zero‑based indices.
        // Start index = 29 (row 30), count = 13 rows (30‑42 inclusive)
        sheet.Cells.DeleteRows(29, 13);

        // Insert five new rows at the position where rows were removed.
        // Insert at index 29 (row 30) with default height.
        sheet.Cells.InsertRows(29, 5);

        // Export the modified workbook to PDF
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
