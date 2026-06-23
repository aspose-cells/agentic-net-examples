using System;
using Aspose.Cells;
using System.Drawing;

class HighlightDuplicates
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the range for column N (zero‑based index 13) rows 0‑100
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 100,
            StartColumn = 13,
            EndColumn = 13
        };

        // Add a new conditional formatting collection
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

        // Assign the range to the conditional formatting
        fcs.AddArea(range);

        // Add a duplicate‑values condition
        int condIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition condition = fcs[condIndex];

        // Set a light orange background for duplicate cells
        condition.Style.BackgroundColor = Color.FromArgb(255, 255, 200, 0);

        // Save the workbook
        workbook.Save("DuplicateHighlight.xlsx");
    }
}