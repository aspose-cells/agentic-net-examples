// Title: Link a textbox shape to a cell formatted with the TEXT function using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link textbox shape to cell value with TEXT formula | display formatted percentage from TEXT function in a shape using Aspose.Cells | set shape text to calculated cell result after workbook.CalculateFormula in Aspose.Cells | C# Aspose.Cells add textbox and bind it to TEXT(A1) result | how to bind a shape to a cell that contains a TEXT function formatted number in Aspose.Cells
// Tags: textbox shape cell binding Aspose.Cells | TEXT formula result shape text Aspose.Cells | calculate workbook formulas Aspose.Cells | format percentage with TEXT function Aspose.Cells | C# shape font styling Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes a numeric value to A1, formats it as a percentage in B1 using the TEXT function, calculates formulas, adds a textbox shape, assigns the shape's text to the formatted string from B1, applies font styling, and saves the file as LinkedShape.xlsx.
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

                // Put a numeric value in cell A1
                sheet.Cells["A1"].PutValue(0.12345);

                // Use TEXT function to format as percentage with two decimals
                sheet.Cells["B1"].Formula = "=TEXT(A1,\"0.00%\")";

                // Calculate formulas so that B1 contains the formatted text
                workbook.CalculateFormula();

                // Add a textbox shape (upperLeftRow, upperLeftColumn, height, width, top, left)
                Shape shape = sheet.Shapes.AddTextBox(5, 5, 30, 200, 5, 5);

                // Link the shape's text to the value of cell B1
                shape.Text = sheet.Cells["B1"].StringValue;

                // Set visual properties for the shape
                shape.Font.Size = 12;
                shape.Font.IsBold = true;
                shape.Fill.FillType = FillType.Solid;
                // Note: SetSolidFillColor is not available in the current API version.
                // The solid fill type is set above; default color will be used.

                // Define output file path
                string outputPath = "LinkedShape.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
