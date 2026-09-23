// Title: How to auto‑fit Excel column widths in C# with Aspose.Cells using pixel‑based measurement of the longest cell text
// AI Prompts: Write C# code that loops through each column of a worksheet, measures the pixel width of the longest string in that column using Aspose.Cells, and sets the column width to that pixel value. | Create a helper method `FitColumnsByPixelWidth(Worksheet sheet)` that computes the maximum text width per column with Aspose.Cells and applies `sheet.SetColumnWidth` accordingly. | Generate a complete example that loads an XLSX file, adjusts all column widths based on the pixel length of the longest cell content, and saves the result while handling missing‑file errors.
// Common Searches: Aspose.Cells C# set column width from pixel measurement of longest cell value | How to calculate pixel width of text in Excel cells using Aspose.Cells .NET | C# auto‑fit Excel columns based on content length in pixels with Aspose.Cells | Adjust column widths programmatically in .xlsx using Aspose.Cells and pixel dimensions
// Tags: Aspose.Cells auto‑fit column width by pixel | C# set Excel column width from content length | Aspose.Cells measure cell text pixel width | Excel .xlsx column width adjustment .NET | programmatic column sizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, accesses the first worksheet, determines the used range, iterates over each column calling `AutoFitColumn` to automatically size columns based on their cell contents, and then saves the modified workbook while handling missing file and generic exceptions.
class AdjustColumnWidths
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxRow;       // zero‑based index of the last used row
            int maxColumn = cells.MaxColumn; // zero‑based index of the last used column

            // Iterate through each column and auto‑fit its width
            for (int col = 0; col <= maxColumn; col++)
            {
                // AutoFitColumn adjusts the column width based on the content of the cells in that column
                sheet.AutoFitColumn(col);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
