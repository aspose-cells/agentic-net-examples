using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied (e.g., A1:A20)
            CellArea range = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                EndRow = 19,       // Row 20
                StartColumn = 0,   // Column A
                EndColumn = 0      // Column A
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the conditional formatting collection
            conditions.AddArea(range);

            // Add a condition: cell value greater than 1000
            int conditionIndex = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "1000",   // Formula1: the threshold value
                null);    // Formula2 not needed for GreaterThan

            // Retrieve the created condition and set its style (green background)
            FormatCondition condition = conditions[conditionIndex];
            condition.Style.BackgroundColor = Color.Green;

            // (Optional) Populate some sample data to demonstrate the rule
            for (int i = 0; i <= 19; i++)
            {
                worksheet.Cells[i, 0].PutValue(i * 200); // Values: 0,200,400,...,3800
            }

            // Save the workbook
            workbook.Save("ConditionalFormatting_GreaterThan1000.xlsx");
        }
    }
}