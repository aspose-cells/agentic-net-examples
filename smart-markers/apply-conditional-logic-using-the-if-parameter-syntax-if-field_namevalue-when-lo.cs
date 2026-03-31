using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook with custom load options
        LoadOptions loadOptions = new LoadOptions();
        // loadOptions.LoadDataOnly = true; // removed because not supported in this version

        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Load the workbook using the specified options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range to which the conditional formatting will be applied (e.g., A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 9,
            EndColumn = 0
        };

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Assign the target range
        fcc.AddArea(area);

        // Add a conditional formatting rule that uses the IF syntax.
        // The formula checks if the cell value equals "Approved" and returns TRUE/FALSE.
        string ifFormula = "=IF(A1=\"Approved\",TRUE,FALSE)";

        // Use an expression type condition; OperatorType.None is appropriate for IF formulas.
        int conditionIdx = fcc.AddCondition(FormatConditionType.Expression, OperatorType.None, ifFormula, null);

        // Optionally style the cells when the condition is met (e.g., green background)
        FormatCondition condition = fcc[conditionIdx];
        condition.Style.BackgroundColor = Color.LightGreen;

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}