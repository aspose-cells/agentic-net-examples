using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate B2:B15 with sample values
        for (int i = 0; i < 14; i++)
        {
            sheet.Cells[i + 1, 1].PutValue((i + 1) * 10); // B column = index 1
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range B2:B15
        CellArea area = new CellArea
        {
            StartRow = 1,   // Row 2 (zero‑based)
            EndRow = 14,    // Row 15
            StartColumn = 1, // Column B (zero‑based)
            EndColumn = 1
        };
        fcs.AddArea(area);

        // Add an IconSet condition
        int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcs[conditionIdx];

        // Set a base icon set type (required before customizing icons)
        condition.IconSet.Type = IconSetType.Arrows3;

        // Enable showing the cell values alongside the icons
        condition.IconSet.ShowValue = true;

        // Customize each icon in the set
        // First icon
        ConditionalFormattingIcon cfIcon0 = condition.IconSet.CfIcons[0];
        cfIcon0.Type = IconSetType.Arrows3;
        cfIcon0.Index = 0;

        // Second icon
        ConditionalFormattingIcon cfIcon1 = condition.IconSet.CfIcons[1];
        cfIcon1.Type = IconSetType.ArrowsGray3;
        cfIcon1.Index = 1;

        // Third icon
        ConditionalFormattingIcon cfIcon2 = condition.IconSet.CfIcons[2];
        cfIcon2.Type = IconSetType.Boxes5;
        cfIcon2.Index = 2;

        // Save the workbook
        workbook.Save("IconSetCustom.xlsx");
    }
}