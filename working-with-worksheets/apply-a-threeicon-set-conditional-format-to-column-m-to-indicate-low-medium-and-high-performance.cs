using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range for column M (zero‑based index 12) from row 0 to row 99
        CellArea columnMRange = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 12,
            EndColumn = 12
        };

        // Add a new conditional formatting entry
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

        // Apply the range to the conditional formatting
        conditions.AddArea(columnMRange);

        // Add an IconSet condition (three‑icon set)
        int conditionIndex = conditions.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = conditions[conditionIndex];

        // Choose a three‑icon set type (e.g., Flags3) and keep the cell values visible
        condition.IconSet.Type = IconSetType.Flags3;
        condition.IconSet.ShowValue = true;

        // Save the workbook
        workbook.Save("ColumnM_IconSet.xlsx");
    }
}