using System;
using Aspose.Cells;
using System.Drawing;

class HighlightDuplicates
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in column B (index 1)
        string[] sampleData = { "Apple", "Banana", "Apple", "Orange", "Banana", "Grape" };
        for (int i = 0; i < sampleData.Length; i++)
        {
            cells[i, 1].PutValue(sampleData[i]); // B column
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

        // Define the range that covers the populated cells in column B
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = sampleData.Length - 1,
            StartColumn = 1,
            EndColumn = 1
        };
        conditions.AddArea(range);

        // Add a condition that highlights duplicate values
        int dupConditionIdx = conditions.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition dupCondition = conditions[dupConditionIdx];

        // Set the visual style for duplicate cells (e.g., yellow background)
        dupCondition.Style.BackgroundColor = Color.Yellow;

        // Save the workbook with the applied conditional formatting
        workbook.Save("DuplicateHighlight.xlsx");
    }
}