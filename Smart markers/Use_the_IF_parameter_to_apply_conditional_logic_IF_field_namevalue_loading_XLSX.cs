using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook (XLSX) from disk
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------- IF LOGIC PARAMETERS --------------------
        // Field (column header) to evaluate
        string fieldName = "Status";
        // Desired value to compare against
        string targetValue = "Approved";

        // Locate the column index of the specified field name in the header row (row 0)
        int fieldColumn = -1;
        for (int col = 0; col <= worksheet.Cells.MaxColumn; col++)
        {
            if (worksheet.Cells[0, col].StringValue.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
            {
                fieldColumn = col;
                break;
            }
        }

        // If the field is not found, exit gracefully
        if (fieldColumn == -1)
        {
            Console.WriteLine($"Field \"{fieldName}\" not found in the header row.");
            return;
        }

        // -------------------- APPLY CONDITIONAL FORMATTING --------------------
        // Define the data range (excluding the header row)
        int firstDataRow = 1;
        int lastDataRow = worksheet.Cells.MaxDataRow;

        // Create a new conditional formatting collection
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Set the area that the conditional format will cover (only the target column)
        CellArea area = new CellArea
        {
            StartRow = firstDataRow,
            EndRow = lastDataRow,
            StartColumn = fieldColumn,
            EndColumn = fieldColumn
        };
        fcc.AddArea(area);

        // Build an expression formula that represents the IF logic:
        //   =($B2="Approved")   (assuming fieldColumn corresponds to column B)
        // Use A1 style for readability
        string columnLetter = CellsHelper.ColumnIndexToName(fieldColumn);
        // Row number in the formula is relative; use the first data row (2 in A1 style)
        string formula = $"=${columnLetter}2=\"{targetValue}\"";

        // Add the expression condition to the collection
        int conditionIdx = fcc.AddCondition(FormatConditionType.Expression, OperatorType.None, formula, null);
        FormatCondition condition = fcc[conditionIdx];

        // Set the visual style applied when the condition is true
        condition.Style.BackgroundColor = Color.LightGreen;

        // -------------------- SAVE RESULT --------------------
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Conditional formatting applied and saved to \"{outputPath}\".");
    }
}