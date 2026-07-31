// Title: Apply Custom Number Format to Negative Formula Results with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, set formulas, calculate them, and apply a red‑currency custom number format only to cells whose evaluated result is negative, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | custom number format | negative values | formula result | StyleFlag | Excel formatting | red negative numbers | programmatic styling | workbook manipulation | .NET
// Common Searches: Aspose.Cells apply custom format to negative results | C# format negative formula values in Excel | red currency format for negative numbers Aspose.Cells | how to style cells based on formula outcome .NET | apply number format only to negative doubles Aspose
// Developer Intent: Apply a custom number format exclusively to cells that contain formulas whose calculated value is negative.
// Use Cases: Highlight negative profit figures in financial statements with a red currency style. | Display inventory adjustments below zero in red while leaving positive values unchanged. | Mark temperature deviations that are below baseline in a scientific worksheet.
// AI Prompts: Generate C# code using Aspose.Cells to apply a red‑currency custom number format only to cells with negative formula results. | Provide an Aspose.Cells solution that uses conditional formatting to color negative formula outcomes red. | Explain how to extend the example to process all worksheets in a workbook while applying the same negative‑result formatting.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, set formulas, calculate them, and apply a red‑currency custom number format only to cells whose evaluated result is negative, using Aspose.Cells for .NET.
    class ApplyCustomNumberFormatToNegativeFormulaResults
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data
                cells["B1"].PutValue(5);   // Positive number
                cells["B2"].PutValue(20);  // Positive number

                // Formulas that may produce negative results
                cells["A1"].SetFormula("=B1-10", null); // Result: -5
                cells["A2"].SetFormula("=B2-15", null); // Result: 5 (positive)

                // Calculate all formulas
                workbook.CalculateFormula();

                // Custom number format (red color for negative numbers)
                string customFormat = "_-\"$\"* #,##0.00_);[Red]-\"$\"* #,##0.00_);_-\"$\"* \"-\"??_);_(@_)";

                // Iterate through used cells and apply the format only to formulas with negative results
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains a formula and its evaluated value is a negative double
                        if (cell.IsFormula && cell.Value != null && cell.Value is double d && d < 0)
                        {
                            // Create a style with the custom number format
                            Style style = workbook.CreateStyle();
                            style.Custom = customFormat;

                            // Apply only the number format using StyleFlag
                            StyleFlag flag = new StyleFlag
                            {
                                NumberFormat = true
                            };

                            // Apply the style to the specific cell
                            Aspose.Cells.Range range = cells.CreateRange(row, col, 1, 1);
                            range.ApplyStyle(style, flag);
                        }
                    }
                }

                // Define output file path
                string outputPath = "NegativeFormulaNumberFormat.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
