// Title: C# – Import a 2D string array into an Aspose.Cells worksheet and enable text wrap
// Description: Creates a workbook, imports a two‑dimensional string array starting at A1, applies a style with text wrapping to every populated cell, auto‑fits rows, and saves the file as WrappedTextOutput.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells import 2D array | C# text wrap cells | StyleFlag.WrapText | AutoFitRows Aspose.Cells | object[,] to worksheet | Excel export long text
// Common Searches: how to import a 2d array into Aspose.Cells | apply text wrap to all cells in Aspose.Cells C# | auto fit rows after wrapping text Aspose.Cells | wrap long strings in Excel with Aspose.Cells
// Developer Intent: Load a two‑dimensional string array into a worksheet and ensure every cell shows wrapped text without truncation.
// Use Cases: Generating Excel reports where description fields contain lengthy paragraphs. | Exporting database query results with multi‑line comments that need automatic row height adjustment. | Building a template that receives dynamic data arrays and formats them for printable output.
// AI Prompts: Write C# code with Aspose.Cells to import a 2D string array and apply text wrapping to all populated cells. | Create a reusable method that takes an object[,] array, imports it, sets IsTextWrapped via StyleFlag, and auto‑fits rows. | Explain how StyleFlag.WrapText works when calling Cells.ApplyStyle in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, imports a two‑dimensional string array starting at A1, applies a style with text wrapping to every populated cell, auto‑fits rows, and saves the file as WrappedTextOutput.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Two‑dimensional array of strings (object[,])
            object[,] data = new object[,]
            {
                {
                    "This is a very long piece of text that should wrap inside cell A1 without being truncated.",
                    "Another long text for cell B1 that also needs wrapping."
                },
                {
                    "Second row, cell A2 with long content to demonstrate wrapping.",
                    "Second row, cell B2 with long content as well."
                }
            };

            // Import the array starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(data, 0, 0);

            // Create a style that enables text wrapping
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Define a style flag to apply only the wrap setting
            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;

            // Apply the wrap style to all cells that contain data
            cells.ApplyStyle(wrapStyle, flag);

            // Auto‑fit rows so the wrapped text becomes visible
            worksheet.AutoFitRows();

            // Save the workbook
            workbook.Save("WrappedTextOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
