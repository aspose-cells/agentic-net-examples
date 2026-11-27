// Title: Automatically highlight Excel cells containing formula errors with Aspose.Cells for .NET using conditional formatting
// Common Searches: Aspose.Cells C# add conditional formatting to flag cells with #DIV/0! errors | How to use ContainsErrors format condition in a .NET workbook | Highlight #N/A and #NAME? cells in Excel using Aspose.Cells API | Apply yellow background and red bold font to error cells programmatically with Aspose.Cells | Create worksheet‑wide error highlighting rule in C# with Aspose.Cells
// Tags: Aspose.Cells conditional formatting for error values | ContainsErrors format condition usage | Excel formula error highlighting with Aspose | style definition for error values Aspose.Cells | worksheet-wide conditional formatting C#

using System;
using System.Drawing;
using Aspose.Cells;

namespace HighlightErrorCells
{
    // The example creates a workbook, inserts values and formulas that generate #DIV/0!, #N/A, and #NAME? errors, calculates all formulas, and then adds a worksheet‑wide conditional formatting rule of type ContainsErrors. The rule applies a yellow background with red bold font to any cell containing an error and saves the file as HighlightedErrors.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate cells with values and formulas that will generate errors
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(0);
                cells["C1"].Formula = "=A1/B1"; // #DIV/0! error

                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(5);
                cells["C2"].Formula = "=VLOOKUP(\"Orange\", A2:B2, 2, FALSE)"; // #N/A error

                cells["A3"].Formula = "=UNKNOWNFUNC(1)"; // #NAME? error

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Add a conditional formatting rule that highlights cells containing errors
                // 1. Create a new conditional formatting collection
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

                // 2. Define the range to which the rule will be applied (entire sheet in this example)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = worksheet.Cells.MaxDataRow,
                    StartColumn = 0,
                    EndColumn = worksheet.Cells.MaxDataColumn
                };
                cfCollection.AddArea(area);

                // 3. Add a condition of type ContainsErrors (highlights cells with formula errors)
                int conditionIndex = cfCollection.AddCondition(FormatConditionType.ContainsErrors);
                FormatCondition condition = cfCollection[conditionIndex];

                // 4. Define the style to apply when the condition is met
                Style style = condition.Style;
                style.BackgroundColor = Color.Yellow;
                style.Font.Color = Color.Red;
                style.Font.IsBold = true;

                // Save the workbook to a file
                string outputPath = "HighlightedErrors.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Cells containing errors are highlighted with a yellow background and red bold font.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
