// Title: Aspose.Cells .NET – Apply Icon Set Conditional Formatting with Text Labels to C3:C12 (C#)
// Description: Creates a new workbook, fills cells C3:C12 with values 1‑10, adds an IconSet conditional formatting rule, switches the icon set to None while keeping ShowValue enabled so the numeric values appear as text, and saves the file as IconSetWithTextLabels.xlsx. Demonstrates how to replace icons with text labels using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | icon set conditional formatting | replace icons with text | ShowValue true | IconSet.Type None | C3:C12 | Excel export sample | GitHub example | conditional formatting code
// Common Searches: Aspose.Cells hide icons in icon set | display values instead of icons Aspose.Cells | icon set conditional formatting C# example | how to set IconSet.Type to None Aspose.Cells | apply conditional formatting to column C in Aspose.Cells
// Developer Intent: Add an IconSet rule to C3:C12, suppress the icons, show only the cell values, and save the workbook.
// Use Cases: Print‑friendly reports where visual cues are needed but icons cannot be rendered. | Excel templates that rely on icon‑set logic for calculations while displaying plain numbers for auditors. | Data exports to environments (e.g., web viewers) that support only text, not graphic icons.
// AI Prompts: Generate C# code using Aspose.Cells to apply an IconSet conditional format to a range and show only the cell values. | Show how to modify an existing IconSet condition so that IconSet.Type is None and ShowValue remains true. | Explain step‑by‑step how to replace icons with text labels in an Aspose.Cells conditional formatting rule for .NET.

using Aspose.Cells;

// Creates a new workbook, fills cells C3:C12 with values 1‑10, adds an IconSet conditional formatting rule, switches the icon set to None while keeping ShowValue enabled so the numeric values appear as text, and saves the file as IconSetWithTextLabels.xlsx. Demonstrates how to replace icons with text labels using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells C3:C12 (rows 2‑11, column 2) with sample numeric values
        for (int row = 2; row <= 11; row++)
        {
            worksheet.Cells[row, 2].PutValue(row - 1); // values 1‑10
        }

        // Add a new conditional formatting entry to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the target range C3:C12
        CellArea area = new CellArea
        {
            StartRow = 2,   // C3
            EndRow = 11,    // C12
            StartColumn = 2,
            EndColumn = 2
        };
        fcc.AddArea(area);

        // Add an IconSet condition to the range
        int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[conditionIndex];

        // Choose any icon set type (e.g., TrafficLights31)
        condition.IconSet.Type = IconSetType.TrafficLights31;

        // Replace the default icons with text labels:
        //  - Ensure the cell values are displayed
        //  - Set the icon set type to None so no icons are shown
        condition.IconSet.ShowValue = true;
        condition.IconSet.Type = IconSetType.None;

        // Save the workbook
        workbook.Save("IconSetWithTextLabels.xlsx");
    }
}
