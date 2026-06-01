using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array of strings to import
        string[,] data = new string[,]
        {
            { "Item", "Description with a very long text that should be wrapped inside the cell to avoid truncation." },
            { "Apple", "A fruit that is typically red, green, or yellow and often eaten fresh or used in recipes." },
            { "Banana", "A long curved fruit that has a thick skin and soft sweet flesh, popular worldwide." }
        };

        // Import the array starting at cell A1 (row 0, column 0)
        cells.ImportTwoDimensionArray(data, 0, 0);

        // Create a style that enables text wrapping
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Define a style flag to apply only the wrap setting
        StyleFlag flag = new StyleFlag();
        flag.WrapText = true;

        // Apply the wrap style to the entire worksheet
        cells.ApplyStyle(wrapStyle, flag);

        // Auto‑fit rows so the wrapped text becomes fully visible
        worksheet.AutoFitRows();

        // Save the workbook
        workbook.Save("WrappedTextOutput.xlsx");
    }
}