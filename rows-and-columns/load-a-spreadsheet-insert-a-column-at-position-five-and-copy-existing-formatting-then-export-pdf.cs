// Title: Insert a column at position five, copy the previous column’s formatting, and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Add a new column at zero‑based index 4 in an existing workbook, duplicate the style from column D for every used row, and save the result as a PDF with Aspose.Cells in C#. | Using Aspose.Cells for .NET, insert a column after the fourth column, copy the formatting from the adjacent column across all rows, then convert the worksheet to a PDF file.
// Common Searches: Aspose.Cells C# insert column at specific index while preserving formatting | How to copy column style when adding a new column with Aspose.Cells .NET | Export Excel workbook to PDF after inserting a column using Aspose.Cells | C# Aspose.Cells keep cell styles after column insertion | Insert column at position 5 and generate PDF with Aspose.Cells
// Tags: insert column Aspose.Cells C# | copy column formatting Aspose.Cells | export workbook to PDF Aspose.Cells | preserve Excel styles Aspose.Cells | column insertion style cloning Aspose.Cells

using System;
using Aspose.Cells;

// Loads input.xlsx, inserts a new column at position five, copies formatting from the preceding column for all populated rows, and saves the modified workbook as output.pdf.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at position five (zero‑based index 4)
        cells.InsertColumn(4);

        // Determine the last used row to limit the formatting copy loop
        int lastRow = cells.MaxDataRow;

        // Copy formatting from the column before the inserted one (column index 3)
        for (int row = 0; row <= lastRow; row++)
        {
            // Get the style from the source cell
            Style sourceStyle = cells[row, 3].GetStyle();

            // Apply the style to the newly inserted column cell
            cells[row, 4].SetStyle(sourceStyle);
        }

        // Export the modified workbook to PDF
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
