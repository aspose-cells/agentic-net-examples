// Title: C# – Apply a Green‑Yellow‑Red Icon Set with Aspose.Cells Conditional Formatting
// Description: Creates a workbook, fills cells A1:A6 with sales figures, defines the range, adds an IconSet conditional format based on the built‑in TrafficLights31 set, customizes the three icons (green, yellow, red), shows values alongside icons, and saves the file as SalesIconSet.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# conditional formatting | icon set | TrafficLights31 | green yellow red icons | sales dashboard | Excel icon set programmatically | Aspose.Cells example
// Common Searches: Aspose.Cells add traffic lights icon set C# | customize icon set conditional formatting Aspose.Cells | apply green yellow red icons to a range with Aspose.Cells | how to use IconSetType in Aspose.Cells .NET
// Developer Intent: Add a three‑color (green, yellow, red) traffic‑light icon set to cells A1:A6 to visually rank sales figures.
// Use Cases: Highlight high, medium, and low sales values with green, yellow, and red icons in a worksheet. | Build a compact dashboard where icons replace numbers for quick performance assessment. | Generate a formatted report that automatically flags low‑selling items in red and top‑selling items in green.
// AI Prompts: Generate C# code using Aspose.Cells to apply a custom green‑yellow‑red icon set to a specified cell range. | Explain how to switch the icon indices to use a different built‑in icon set in Aspose.Cells conditional formatting. | Show the steps to export a workbook containing conditional icon sets to PDF while preserving the icons.

using Aspose.Cells;

// Creates a workbook, fills cells A1:A6 with sales figures, defines the range, adds an IconSet conditional format based on the built‑in TrafficLights31 set, customizes the three icons (green, yellow, red), shows values alongside icons, and saves the file as SalesIconSet.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample sales figures in column A (A1:A6)
        double[] sales = { 1200, 800, 500, 1500, 300, 950 };
        for (int i = 0; i < sales.Length; i++)
        {
            worksheet.Cells[i, 0].PutValue(sales[i]);
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A6 for the conditional formatting
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

        // Use the built‑in TrafficLights31 set as the base (green, yellow, red)
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true;   // display cell values alongside icons
        condition.IconSet.Reverse = false;    // keep default order

        // Customize individual icons (green, yellow, red)
        // Index 0 = green, 1 = yellow, 2 = red in TrafficLights31
        ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
        icon0.Type = IconSetType.TrafficLights31;
        icon0.Index = 0; // green

        ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
        icon1.Type = IconSetType.TrafficLights31;
        icon1.Index = 1; // yellow

        ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
        icon2.Type = IconSetType.TrafficLights31;
        icon2.Index = 2; // red

        // Save the workbook with the applied conditional icon set
        workbook.Save("SalesIconSet.xlsx");
    }
}
