// Title: C# – Apply Traffic Lights Icon Set Conditional Formatting and Save Workbook as XLSM using Aspose.Cells
// Description: Loads an existing workbook, enables macro support, applies a TrafficLights31 icon set to cells A1:A10 with values displayed, and saves the file as a macro‑enabled XLSM.
// Keywords: Aspose.Cells C# icon set | conditional formatting traffic lights | show values with icon set | save as XLSM Aspose | enable macros Aspose.Cells | C# Excel conditional formatting | macro enabled workbook .NET | Aspose.Cells SaveFormat.Xlsm
// Common Searches: Aspose.Cells add traffic lights icon set C# | How to enable macros and save as XLSM with Aspose.Cells | Display cell values with icon set in Aspose.Cells .NET | Conditional formatting icon set range A1:A10 Aspose | Save workbook as macro enabled file using Aspose.Cells
// Developer Intent: Add an icon set that shows cell values and export the workbook as a macro‑enabled XLSM file.
// Use Cases: Create performance dashboards where icons highlight status while numeric values remain visible. | Prepare Excel reports that must retain existing VBA macros after formatting. | Automate generation of XLSM files with conditional icon formatting for distribution to end‑users.
// AI Prompts: Write C# code with Aspose.Cells to apply a 3‑color traffic lights icon set to range B2:B20, show values, and save as XLSM. | Explain how to change the icon set type and configure thresholds in Aspose.Cells conditional formatting. | Show how to enable macros and preserve existing VBA when saving a workbook after adding conditional formatting with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an existing workbook, enables macro support, applies a TrafficLights31 icon set to cells A1:A10 with values displayed, and saves the file as a macro‑enabled XLSM.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook wb = new Workbook("input.xlsx");

        // Enable macros so the workbook can be saved as XLSM
        wb.Settings.EnableMacros = true;

        // Work with the first worksheet
        Worksheet sheet = wb.Worksheets[0];

        // Define the cell range that will receive the icon set (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Get the conditional formatting collection of the worksheet
        ConditionalFormattingCollection cfs = sheet.ConditionalFormattings;

        // Add a new conditional formatting rule and set its range
        int cfIndex = cfs.Add();
        FormatConditionCollection fcc = cfs[cfIndex];
        fcc.AddArea(area);

        // Add an IconSet condition to the rule
        int condIndex = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[condIndex];

        // Configure the icon set: choose a type and ensure cell values are shown
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true; // display the cell values alongside icons

        // Save the workbook as a macro‑enabled XLSM file
        wb.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
