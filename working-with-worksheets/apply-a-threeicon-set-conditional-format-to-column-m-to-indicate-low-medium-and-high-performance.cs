// Title: C# – Apply Traffic Lights Icon Set Conditional Formatting to Column M using Aspose.Cells
// Description: Shows how to create a workbook, define a CellArea for column M (rows 0‑99), add an IconSet conditional format of type TrafficLights31, set 33 % and 67 % thresholds, display cell values, populate sample data, and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# conditional formatting | icon set | traffic lights icon set | column M formatting | percentile thresholds | Excel conditional formatting .NET | FormatConditionType.IconSet | IconSetType.TrafficLights31 | CellArea | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add traffic lights icon set to a column | C# conditional formatting icon set percent thresholds | How to set three‑icon conditional format in Aspose.Cells | Apply icon set conditional formatting to column M in .NET | Aspose.Cells IconSet example with percent values
// Developer Intent: Add a three‑icon Traffic Lights conditional format to column M to represent low, medium, and high performance values.
// Use Cases: Generate performance dashboards where numeric scores are visualized with red, yellow, and green icons. | Create Excel templates that automatically color‑code sales or KPI columns based on percentile ranges. | Automate test reports that include sample data and icon‑based visual cues for quick assessment. | Export data from an application and apply icon set formatting before delivering the workbook to end users.
// AI Prompts: Provide C# code to replace the Traffic Lights icon set with a 4‑arrow set and set thresholds at 20 % and 80 % using Aspose.Cells. | Show how to hide the cell values while keeping only the icons visible in the conditional format. | Explain how to copy the IconSet conditional formatting from column M to another column programmatically. | Give an example of applying the same icon set to a dynamic range that expands beyond row 99. | Describe how to customize the icon order (reverse) for the Traffic Lights set.

using Aspose.Cells;

// Shows how to create a workbook, define a CellArea for column M (rows 0‑99), add an IconSet conditional format of type TrafficLights31, set 33 % and 67 % thresholds, display cell values, populate sample data, and save the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range for column M (zero‑based index 12), rows 0‑99
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 12,
            EndColumn = 12
        };

        // Add a conditional formatting collection and set its range
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];
        fcs.AddArea(area);

        // Add an IconSet condition to the collection
        int condIdx = fcs.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcs[condIdx];

        // Configure the icon set (Traffic Lights: red, yellow, green)
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true;   // display cell values alongside icons
        condition.IconSet.Reverse = false;    // keep default icon order

        // Define three thresholds: low (≤33%), medium (≤67%), high (>67%)
        condition.IconSet.Cfvos[0].Type = FormatConditionValueType.Percent;
        condition.IconSet.Cfvos[0].Value = 33;
        condition.IconSet.Cfvos[0].IsGTE = true;   // use greater‑than‑or‑equal

        condition.IconSet.Cfvos[1].Type = FormatConditionValueType.Percent;
        condition.IconSet.Cfvos[1].Value = 67;
        condition.IconSet.Cfvos[1].IsGTE = true;   // use greater‑than‑or‑equal

        // The third CFVO is automatically Max; no changes needed

        // (Optional) Populate sample data in column M for demonstration
        for (int i = 0; i < 100; i++)
        {
            sheet.Cells[i, 12].PutValue(i); // values 0‑99
        }

        // Save the workbook
        workbook.Save("ColumnM_IconSet.xlsx", SaveFormat.Xlsx);
    }
}
