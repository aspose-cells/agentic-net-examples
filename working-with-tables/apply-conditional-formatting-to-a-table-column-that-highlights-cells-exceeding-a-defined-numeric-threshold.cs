using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column B (index 1) rows 1-20
            for (int row = 0; row < 20; row++)
            {
                sheet.Cells[row, 1].PutValue(row * 5); // Example values: 0,5,10,...
            }

            // Define the range for the column to which the conditional format will be applied
            // Here we target column B from row 0 to row 19 (A1 style: B1:B20)
            CellArea columnArea = new CellArea
            {
                StartRow = 0,
                EndRow = 19,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the conditional formatting collection
            conditions.AddArea(columnArea);

            // Define the numeric threshold; cells greater than this value will be highlighted
            const string threshold = "50";

            // Add a CellValue condition: GreaterThan the threshold
            int conditionIdx = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold,
                null);

            // Retrieve the created condition and set its formatting style
            FormatCondition condition = conditions[conditionIdx];
            condition.Style.BackgroundColor = Color.Yellow; // Highlight with yellow background
            condition.Style.Font.Color = Color.Black;       // Optional: set font color for readability

            // Save the workbook to a file
            workbook.Save("ConditionalFormatting_ColumnThreshold.xlsx");
        }
    }
}