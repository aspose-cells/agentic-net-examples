using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column F (index 5)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 5].PutValue(i * 10 + 5);
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for column F (rows 0‑9)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 5,
            EndColumn = 5
        };
        fcc.AddArea(area);

        // Add an AboveAverage condition to the collection
        int conditionIndex = fcc.AddCondition(FormatConditionType.AboveAverage);
        FormatCondition fc = fcc[conditionIndex];

        // Configure the condition to highlight cells above the average
        fc.AboveAverage.IsAboveAverage = true;
        fc.Style.BackgroundColor = Color.Yellow;

        // Optional: autofit columns for better visibility
        worksheet.AutoFitColumns();

        // Save the workbook
        workbook.Save("ColumnF_AboveAverage.xlsx");
    }
}