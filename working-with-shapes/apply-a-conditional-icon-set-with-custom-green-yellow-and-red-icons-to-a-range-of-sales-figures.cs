using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample sales figures in column A (A1:A5)
        double[] sales = { 5000, 12000, 8000, 15000, 3000 };
        for (int i = 0; i < sales.Length; i++)
        {
            worksheet.Cells[i, 0].PutValue(sales[i]);
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A5 for the conditional formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = sales.Length - 1,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add an IconSet condition
        int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[conditionIndex];

        // Use the built‑in TrafficLights31 icon set (green, yellow, red)
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true;   // display cell values alongside icons
        condition.IconSet.Reverse = false;    // keep default order

        // Customize individual icons to ensure the order is Green → Yellow → Red
        // Icon 0 (lowest value) – Green
        ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
        icon0.Type = IconSetType.TrafficLights31;
        icon0.Index = 2; // index 2 corresponds to the green light

        // Icon 1 (middle value) – Yellow
        ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
        icon1.Type = IconSetType.TrafficLights31;
        icon1.Index = 1; // index 1 corresponds to the yellow light

        // Icon 2 (highest value) – Red
        ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
        icon2.Type = IconSetType.TrafficLights31;
        icon2.Index = 0; // index 0 corresponds to the red light

        // Save the workbook with the applied conditional icon set
        workbook.Save("SalesIconSet.xlsx");
    }
}