// Title: Import a 2D string array into an Aspose.Cells worksheet and apply text wrapping to every cell using C#
// AI Prompts: Load a two‑dimensional string array into a worksheet with Worksheet.Cells.ImportTwoDimensionArray, then create a Style with IsTextWrapped = true and assign it to the worksheet's used range. | Generate a wrap‑enabled style in Aspose.Cells, iterate over the MaxDisplayRange after importing data, and set each cell's style to enable text wrapping. | Save the workbook after applying the wrap style to ensure long text is displayed without truncation in the resulting .xlsx file.
// Common Searches: C# Aspose.Cells import 2D string array and wrap text in cells | how to enable text wrapping for all cells after importing data with Aspose.Cells | Aspose.Cells set IsTextWrapped property for a range of cells in .NET | prevent text truncation in Excel by applying wrap style using Aspose.Cells C# | apply a style to the used range after ImportTwoDimensionArray in Aspose.Cells
// Tags: import two-dimensional array Aspose.Cells C# | text wrap style for used range | IsTextWrapped property Aspose.Cells | wrap long text Excel cells C# | save formatted workbook Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new Workbook, imports a two‑dimensional string array into the first worksheet, builds a Style with text wrapping enabled, applies that style to every cell in the used range, and saves the file as WrappedArray.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Sample two‑dimensional array of strings
            string[,] data = new string[,]
            {
                { "Header1", "Header2", "Header3" },
                { "This is a long piece of text that should wrap within the cell", "Short", "Another long text that needs wrapping" },
                { "Row3Col1", "Row3Col2", "Row3Col3" }
            };

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Import the array starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportTwoDimensionArray(data, 0, 0);

            // Get the used range of cells after import
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Create a style with text wrap enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Apply the wrap style to every cell in the used range
            foreach (Cell cell in usedRange)
            {
                cell.SetStyle(wrapStyle);
            }

            // Save the workbook to a file
            workbook.Save("WrappedArray.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
