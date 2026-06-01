using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with text values that will be represented by emojis
            sheet.Cells["A1"].PutValue("Excellent");
            sheet.Cells["A2"].PutValue("Good");
            sheet.Cells["A3"].PutValue("Poor");

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A3 for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add an IconSet condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcc[conditionIdx];

            // Hide the cell text (show only icons)
            condition.IconSet.ShowValue = false;

            // Add custom icons (emoji‑like placeholders)
            // Icon 0 – representing "Excellent"
            condition.IconSet.CfIcons.Add(IconSetType.Smilies3, 0);
            // Icon 1 – representing "Good"
            condition.IconSet.CfIcons.Add(IconSetType.Symbols3, 1);
            // Icon 2 – representing "Poor"
            condition.IconSet.CfIcons.Add(IconSetType.Arrows3, 2);

            // Save the workbook
            workbook.Save("EmojiIconSet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}