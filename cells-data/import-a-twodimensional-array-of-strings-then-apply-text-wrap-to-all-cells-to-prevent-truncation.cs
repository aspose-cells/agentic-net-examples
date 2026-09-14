// Title: How to import a 2D string array into an Aspose.Cells worksheet and enable text wrap for all cells in C#
// AI Prompts: Load a two‑dimensional string array into the first worksheet of a new Workbook and apply a wrap‑text style to every cell using Aspose.Cells for .NET. | Create a Style with IsTextWrapped set to true, configure a StyleFlag that only affects WrapText, and apply it to the worksheet’s Cells collection. | Auto‑fit the rows after the wrap style is applied and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# import 2D string array and wrap text in all cells | apply text wrapping to entire worksheet after ImportTwoDimensionArray Aspose.Cells | auto fit rows after enabling wrap text with Aspose.Cells .NET example
// Tags: import 2d string array wrap text Aspose.Cells | StyleFlag wrap text Aspose.Cells | auto fit rows after wrap text Aspose.Cells | apply style to all worksheet cells Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, imports a two‑dimensional string array starting at A1, defines a style with IsTextWrapped enabled, applies the style to every cell using a StyleFlag that only wraps text, auto‑fits the rows so the wrapped content is fully visible, and saves the file as WrappedTextOutput.xlsx.
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
            { "This is a very long piece of text that should wrap inside the cell to avoid truncation.", "Short" },
            { "Another lengthy text entry that needs wrapping for proper display.", "Medium length text" }
        };

        // Import the array starting at cell A1 (row 0, column 0)
        cells.ImportTwoDimensionArray(data, 0, 0);

        // Create a style with text wrapping enabled
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Create a style flag to apply only the wrap setting
        StyleFlag flag = new StyleFlag();
        flag.WrapText = true;

        // Apply the wrap style to all cells in the worksheet
        cells.ApplyStyle(wrapStyle, flag);

        // Auto‑fit rows so the wrapped text becomes fully visible
        worksheet.AutoFitRows();

        // Save the workbook
        workbook.Save("WrappedTextOutput.xlsx");
    }
}
