// Title: Insert a new column at position 2, copy its style from the left column, and export the worksheet to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert column at second position and copy formatting | How to duplicate column styles after inserting a column using Aspose.Cells .NET | Save modified Excel workbook as PDF with Aspose.Cells in C# | Copy cell style from one column to another before PDF export using Aspose.Cells | Insert column with style copy and export to PDF Aspose.Cells example
// Tags: insert column with style copy Aspose.Cells | copy column cell style C# Aspose.Cells | save worksheet as PDF Aspose.Cells .NET | Excel to PDF conversion after column insertion Aspose.Cells | preserve formatting when inserting column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing Excel file, checks its presence, inserts a new column at index 1 (second position) while copying the style from the adjacent left column, iterates through all populated rows to transfer the original column's cell styles to the new column, ensures the output directory exists, and finally saves the updated workbook as a PDF file using Aspose.Cells' SaveFormat.Pdf, with basic error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a new column at position two (zero‑based index 1)
            // The second argument 'true' copies the style from the left column
            sheet.Cells.InsertColumn(1, true);

            // Copy cell styles from the original column (now at index 0) to the new column (index 1)
            // Iterate through all rows that have data
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Source cell (original column)
                Cell sourceCell = sheet.Cells[row, 0];
                // Destination cell (new column)
                Cell destCell = sheet.Cells[row, 1];

                // Copy the style
                destCell.SetStyle(sourceCell.GetStyle());
            }

            // Ensure the output directory exists
            string outputPath = "output.pdf";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
