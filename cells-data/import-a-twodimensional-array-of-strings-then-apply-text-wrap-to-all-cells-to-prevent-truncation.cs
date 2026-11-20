// Title: Import a 2D string array into an Excel worksheet and apply text wrap to all cells using Aspose.Cells for .NET
// Common Searches: how to import a two dimensional string array into an Excel sheet using Aspose.Cells C# | apply text wrap to all cells after importing data with Aspose.Cells | prevent cell text truncation when using ImportTwoDimensionArray in Aspose.Cells | set wrap text style for a range in Aspose.Cells .NET
// Tags: ImportTwoDimensionArray with wrap style Aspose.Cells | apply wrap text to range Aspose.Cells | prevent cell text truncation Aspose.Cells | style flag wrap text Excel .NET | save workbook as Xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, imports a two‑dimensional string array starting at A1, defines a style with text wrapping enabled, applies the wrap style to the entire imported range using a StyleFlag, and saves the file as WrappedText.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Define a two‑dimensional array of strings
            string[,] data = new string[,]
            {
                { "Header1", "Header2", "Header3" },
                { "This is a long piece of text that should wrap inside the cell", "Short", "Another long text that needs wrapping" },
                { "Row3Col1", "Row3Col2", "Row3Col3" }
            };

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Import the array into the worksheet starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportTwoDimensionArray(data, 0, 0);

            // Determine the size of the imported range
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            // Create a style with text wrap enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Define which style attributes to apply (only wrap text)
            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;

            // Apply the wrap style to the entire imported range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, rows, cols);
            range.ApplyStyle(wrapStyle, flag);

            // Define output file path
            string outputPath = "WrappedText.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
