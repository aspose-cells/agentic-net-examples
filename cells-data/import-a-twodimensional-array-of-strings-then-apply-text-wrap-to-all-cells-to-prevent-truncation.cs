// Title: C# – Import a 2D String Array and Enable Wrap Text for All Cells with Aspose.Cells
// Description: Demonstrates how to load a string[,] into a new workbook, convert it to object[,] for ImportTwoDimensionArray, apply a wrap‑text style across the used range, auto‑fit rows, and save the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# import two dimensional array | wrap text all cells Aspose.Cells | auto fit rows after wrapping | StyleFlag WrapText .NET | Excel text wrapping programmatically | global developers | US .NET Excel library
// Common Searches: how to import a 2d string array into Aspose.Cells | apply wrap‑text to an entire worksheet in C# | auto‑fit rows after setting IsTextWrapped | Aspose.Cells ImportTwoDimensionArray example | C# code to prevent Excel cell truncation
// Developer Intent: Load a multidimensional string array into a worksheet and ensure every cell displays wrapped text without being cut off.
// Use Cases: Populate a report table with long descriptions and keep all text visible. | Generate Excel invoices where comment fields contain extensive notes. | Create a data export from a database array while preserving readability of wrapped content.
// AI Prompts: Provide C# code that imports a string[,] into an Aspose.Cells worksheet and applies a wrap‑text style to the whole used range. | Show how to combine Style, StyleFlag, and AutoFitRows after ImportTwoDimensionArray to avoid text truncation. | Write an example that converts a string matrix to object[,] for Aspose.Cells, enables text wrapping for every cell, and saves the workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsWrapExample
{
    // Demonstrates how to load a string[,] into a new workbook, convert it to object[,] for ImportTwoDimensionArray, apply a wrap‑text style across the used range, auto‑fit rows, and save the result as an Excel file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // 2. Prepare a two‑dimensional array of strings
                string[,] data = new string[,]
                {
                    { "Header 1", "Header 2", "Header 3" },
                    { "Short", "This is a very long piece of text that should wrap inside the cell to avoid truncation.", "Another long text that needs wrapping as well." },
                    { "Row3Col1", "Row3Col2", "Row3Col3" }
                };

                // 3. Convert the string[,] to object[,] required by ImportTwoDimensionArray
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                object[,] objData = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        objData[r, c] = data[r, c];
                    }
                }

                // 4. Import the data starting at cell A1 (row 0, column 0)
                cells.ImportTwoDimensionArray(objData, 0, 0);

                // 5. Create a style with text wrapping enabled
                Style wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true;

                // 6. Create a StyleFlag that applies only the wrap setting
                StyleFlag flag = new StyleFlag();
                flag.WrapText = true;

                // 7. Apply the wrap style to the whole used range
                AsposeRange usedRange = cells.CreateRange(0, 0, rows, cols);
                usedRange.ApplyStyle(wrapStyle, flag);

                // 8. Auto‑fit rows so the wrapped text becomes visible
                worksheet.AutoFitRows();

                // 9. Save the workbook
                string outputPath = "WrappedTextOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
