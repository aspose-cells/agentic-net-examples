// Title: Aspose.Cells .NET: Auto‑highlight formula errors using conditional formatting
// Description: Demonstrates how to create a workbook, add formulas that generate errors, calculate them, and apply a ContainsErrors conditional formatting rule that colors error cells (e.g., yellow) before saving the file.
// Keywords: Aspose.Cells conditional formatting | C# highlight error cells | ContainsErrors format condition | auto highlight formula errors | Aspose.Cells CalculateFormula | .NET spreadsheet error detection
// Common Searches: Aspose.Cells highlight #DIV/0! error C# | Conditional formatting for formula errors Aspose.Cells | Programmatically detect and color error cells in .NET workbook | Apply ContainsErrors rule with Aspose.Cells
// Developer Intent: Add a conditional formatting rule that automatically colors any cell containing a formula error after calculation.
// Use Cases: Mark division‑by‑zero or undefined function errors in financial models. | Flag invalid references in data‑validation sheets before distribution. | Provide instant visual cues for cells that failed to evaluate during bulk calculations.
// AI Prompts: Generate C# code with Aspose.Cells that applies a red font to cells containing any formula error after workbook.CalculateFormula(). | Show how to create two conditional formatting rules in the same worksheet: one for errors (yellow background) and one for negative numbers (red background). | Explain how to retrieve the addresses of cells that triggered the ContainsErrors condition after calculation using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsErrorHighlight
{
    // Demonstrates how to create a workbook, add formulas that generate errors, calculate them, and apply a ContainsErrors conditional formatting rule that colors error cells (e.g., yellow) before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add formulas that will produce errors
            cells["A1"].Formula = "=1/0";                 // Division by zero error
            cells["A2"].Formula = "=UNKNOWNFUNC()";       // Unknown function error
            cells["A3"].Formula = "=B1+1";                // No error (valid reference)

            // Calculate all formulas so that error values are generated
            workbook.CalculateFormula();

            // Define a conditional formatting rule that highlights cells containing errors
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Apply the rule to a range (e.g., A1:E10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 4
            };
            cfCollection.AddArea(area);

            // Add the "ContainsErrors" condition
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.ContainsErrors);
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the highlight style (yellow background)
            condition.Style.BackgroundColor = Color.Yellow;

            // Save the workbook
            workbook.Save("ErrorHighlightDemo.xlsx");
        }
    }
}
