// Title: Using Aspose.Cells Range to populate a matrix layout from a 2D string array in C#
// Common Searches: c# aspnet fill excel range with two dimensional string array using aspose.cells | how to create a matrix layout starting at B2 with Aspose.Cells Range | apply bold gray header style to first row of a range in Aspose.Cells C# | populate excel worksheet from multidimensional array with Aspose.Cells | save workbook as xlsx after filling range with array Aspose.Cells C#
// Tags: populate range from two‑dimensional array Aspose.Cells | matrix layout creation with Range API C# | apply bold light‑gray header style Aspose.Cells | initialize range at B2 cell Aspose.Cells | export workbook to XLSX using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMatrixFill
{
    // The example creates a new Workbook, defines a 4×3 string[,] array, builds a Range starting at cell B2 sized to the array, assigns the array to the Range's Value, styles the first row as a bold light‑gray header, and saves the result as MatrixLayout.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define a two‑dimensional source array
                string[,] sourceData = new string[,]
                {
                    { "Header1", "Header2", "Header3" },
                    { "Row1Col1", "Row1Col2", "Row1Col3" },
                    { "Row2Col1", "Row2Col2", "Row2Col3" },
                    { "Row3Col1", "Row3Col2", "Row3Col3" }
                };

                // Determine the start cell for the matrix layout (e.g., B2)
                int startRow = 1;      // zero‑based index (row 2)
                int startColumn = 1;   // zero‑based index (column B)

                // Create a range that matches the size of the source array
                int rows = sourceData.GetLength(0);
                int columns = sourceData.GetLength(1);
                AsposeRange targetRange = sheet.Cells.CreateRange(startRow, startColumn, rows, columns);

                // Fill the range with the source data
                targetRange.Value = sourceData;

                // Optionally, apply a style to the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the first row of the range (the header)
                AsposeRange headerRange = sheet.Cells.CreateRange(startRow, startColumn, 1, columns);
                headerRange.ApplyStyle(headerStyle, new StyleFlag() { FontBold = true, CellShading = true });

                // Save the workbook to a file
                string outputPath = "MatrixLayout.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
