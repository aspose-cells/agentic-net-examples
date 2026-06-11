using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range for column F (zero‑based index 5), rows 0‑99
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the defined range with the collection
        fcc.AddArea(area);

        // Add a condition: highlight cells with values greater than 1000
        int condIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "1000", null);
        FormatCondition fc = fcc[condIndex];

        // Set the formatting style (light yellow background)
        fc.Style.BackgroundColor = Color.LightYellow;

        // Save the workbook
        workbook.Save("ConditionalFormatting_ColumnF.xlsx");
    }
}