// Title: Automatically highlight formula error cells in an Excel workbook using Aspose.Cells for .NET conditional formatting
// AI Prompts: Write C# code that creates a workbook, inserts formulas, calls CalculateFormula, and adds a ContainsErrors conditional formatting rule that colors error cells with a light salmon background. | Adapt the sample to apply the ContainsErrors conditional formatting only to the range A1:B10 after the workbook has been calculated.
// Common Searches: Aspose.Cells C# highlight cells that contain #DIV/0! after calculating formulas | How to add a ContainsErrors conditional format to an Excel file with Aspose.Cells | Programmatically color Excel error values using Aspose.Cells .NET API | Apply conditional formatting for formula errors only in a specific range with Aspose.Cells | Auto style error cells after workbook.CalculateFormula in C#
// Tags: Aspose.Cells ContainsErrors conditional formatting | C# highlight Excel formula errors | apply background color to error cells Aspose.Cells | calculate formulas then style error values | conditional formatting for #DIV/0! Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace HighlightErrorCellsDemo
{
    // The example creates a workbook, inserts formulas that generate errors, calculates all formulas, and then adds a ContainsErrors conditional formatting rule that fills any error cells with a light salmon background before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some cells with formulas that will produce errors
            cells["A1"].Formula = "=1/0";                 // #DIV/0! error
            cells["A2"].Formula = "=UNKNOWNFUNC(5)";      // #NAME? error
            cells["A3"].Formula = "=B1";                  // B1 is empty, no error (valid)
            cells["B1"].PutValue(10);
            cells["B2"].Formula = "=A2+A3";               // Depends on error cell A2

            // Calculate all formulas so that error values are materialized
            workbook.CalculateFormula();

            // Add a conditional formatting rule that highlights cells containing errors
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the rule applies (entire used range)
            CellArea area = CellArea.CreateCellArea(0, 0, cells.MaxDataRow, cells.MaxDataColumn);
            cfCollection.AddArea(area);

            // Add a condition of type ContainsErrors (no operator or formulas needed)
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.ContainsErrors);
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the style for cells that meet the condition (e.g., light red background)
            condition.Style.BackgroundColor = Color.LightSalmon;

            // Save the workbook
            workbook.Save("HighlightedErrors.xlsx");
        }
    }
}
