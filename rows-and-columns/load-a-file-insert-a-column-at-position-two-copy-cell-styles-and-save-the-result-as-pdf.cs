// Title: Insert a new column at position 2, copy original column styles, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a column at index 1, copies the style from column A to the new column for every row, and saves the workbook as a PDF file. | Generate a .NET routine that loads an Excel workbook, adds a second column while preserving the formatting of the first column, and converts the updated sheet to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# insert column at second position and keep original formatting | How to copy column A style to a newly inserted column with Aspose.Cells .NET | Save Excel workbook as PDF after adding a column using Aspose.Cells | C# Aspose.Cells duplicate cell style between columns before PDF export | Preserve formatting when inserting a column in Aspose.Cells and convert to PDF
// Tags: insert column Aspose.Cells C# | duplicate column formatting Aspose.Cells | export workbook to PDF Aspose.Cells | preserve cell style after column insertion | copy cell style between columns Aspose.Cells

using Aspose.Cells;
using System;

// Loads input.xlsx, inserts a column at index 1, copies the style from column A to the new column for all rows, and saves the workbook as output.pdf in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new column at position two (zero‑based index 1)
        sheet.Cells.InsertColumn(1);

        // Copy cell styles from the original first column (index 0) to the new column (index 1)
        int maxRow = sheet.Cells.MaxDataRow;
        int totalRows = Math.Max(maxRow + 1, sheet.Cells.MaxRow + 1);
        for (int row = 0; row < totalRows; row++)
        {
            Cell sourceCell = sheet.Cells[row, 0];
            Cell destCell = sheet.Cells[row, 1];
            destCell.SetStyle(sourceCell.GetStyle());
        }

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
