// Title: Use Aspose.Cells Range to populate a matrix layout from a 2D object array in C#
// AI Prompts: Generate C# code that creates a Workbook, defines a Range sized to a two‑dimensional object[,] and assigns the array to the Range's Value property with Aspose.Cells. | Provide a C# snippet that formats the first row of the filled Range with bold text and a light gray background using Aspose.Cells styling APIs. | Write C# code to persist the workbook containing the populated matrix range to an Excel file named MatrixLayout.xlsx via Aspose.Cells.
// Common Searches: how to fill an Excel range with a 2d object array using Aspose.Cells C# | Aspose.Cells create matrix layout from object[,] example | apply header formatting to first row of a range in Aspose.Cells | save workbook after populating range Aspose.Cells C# | determine range size from two dimensional array Aspose.Cells
// Tags: populate range from 2d object array Aspose.Cells | define matrix range dimensions C# | header row styling Aspose.Cells | save workbook as Excel file Aspose.Cells | matrix layout using Aspose.Cells Range

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Demonstrates creating a workbook, defining a Range that matches a 2D object array, assigning the array to the Range, styling the header row with bold font and gray background, and saving the result as MatrixLayout.xlsx using Aspose.Cells for .NET.
class MatrixLayoutExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Two‑dimensional source array.
            object[,] sourceData = new object[,]
            {
                { "Item", "Qty", "Price" },
                { "Apple", 10, 0.5 },
                { "Banana", 5, 0.3 },
                { "Cherry", 20, 0.2 }
            };

            // Determine the size of the source array.
            int rows = sourceData.GetLength(0);
            int cols = sourceData.GetLength(1);

            // Define a range where the matrix will be placed (starting at cell A1).
            // Parameters: startRow, startColumn, totalRows, totalColumns.
            AsposeRange matrixRange = sheet.Cells.CreateRange(0, 0, rows, cols);

            // Fill the defined range with the two‑dimensional array.
            matrixRange.Value = sourceData;

            // Optionally, apply a simple style to the header row.
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the first row of the range (header).
            for (int c = 0; c < cols; c++)
            {
                Cell headerCell = sheet.Cells[0, c];
                headerCell.SetStyle(headerStyle);
            }

            // Save the workbook to a file.
            workbook.Save("MatrixLayout.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
