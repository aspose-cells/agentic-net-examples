// Title: Apply red fill conditional formatting to cells with values greater than 100 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that scans the used range of a worksheet and adds a red background conditional formatting rule for any cell whose numeric value exceeds 100. | Create a function in C# using Aspose.Cells that defines a CellValue > 100 condition, applies a solid red style, and attaches the rule to each matching cell in a workbook.
// Common Searches: Aspose.Cells C# apply conditional formatting to cells greater than 100 | how to set red background for numeric values over 100 with Aspose.Cells .NET | programmatically add CellValue > 100 conditional format in an Excel workbook using Aspose.Cells | iterate over used range and apply conditional formatting rule Aspose.Cells C# | Aspose.Cells conditional formatting based on numeric threshold example
// Tags: Aspose.Cells conditional formatting numeric threshold | C# iterate worksheet cells apply style | Aspose.Cells red fill style for values > 100 | CellArea single cell conditional rule Aspose.Cells | CreateStyle background color Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads (or creates) an Excel workbook, iterates through every used cell, and for each numeric cell with a value greater than 100 adds a conditional formatting rule that fills the cell with a solid red background, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                // Populate some sample data
                tempSheet.Cells["A1"].PutValue(50);
                tempSheet.Cells["B2"].PutValue(150);
                tempSheet.Cells["C3"].PutValue(200);
                tempWb.Save(inputPath);
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var sheet = workbook.Worksheets[0];

            // Iterate through all used cells and apply conditional formatting to values > 100
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];
                    if (double.TryParse(cell.StringValue, out double numericValue) && numericValue > 100)
                    {
                        // Add a new conditional formatting rule for this cell
                        int cfIndex = sheet.ConditionalFormattings.Add();
                        var cf = sheet.ConditionalFormattings[cfIndex];

                        // Define the area for this single cell
                        var area = new CellArea
                        {
                            StartRow = row,
                            StartColumn = col,
                            EndRow = row,
                            EndColumn = col
                        };
                        cf.AddArea(area);

                        // Define the condition (cell value > 100)
                        int conditionIndex = cf.AddCondition(
                            FormatConditionType.CellValue,
                            OperatorType.GreaterThan,
                            "100",
                            string.Empty); // formula2 not required for this operator

                        // Configure the style to apply when the condition is met
                        var style = workbook.CreateStyle();
                        style.ForegroundColor = Color.Red;
                        style.Pattern = BackgroundType.Solid;

                        cf[conditionIndex].Style = style;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
