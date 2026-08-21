// Title: C# – Apply Traffic Lights Icon Set Conditional Formatting and Export Workbook as XLSM Using Aspose.Cells
// Description: Loads an XLSX file, activates macro support, creates a conditional formatting rule that applies the TrafficLights31 icon set to cells A1:A10 while keeping the numeric values visible, and saves the result as a macro‑enabled XLSM document.
// Keywords: Aspose.Cells | C# Excel automation | icon set conditional formatting | TrafficLights31 | show cell values | macro enabled workbook | SaveFormat.Xlsm | Excel conditional icons | enable macros Aspose
// Common Searches: asp.net add traffic lights icon set Aspose.Cells | save workbook as xlsm with macros Aspose | display values with icon set conditional formatting | enable macros in Aspose.Cells workbook | apply conditional formatting to range A1:A10 C#
// Developer Intent: Create a workbook that uses an icon‑set rule (Traffic Lights) with visible numbers and write it out as an XLSM file that preserves macro capability.
// Use Cases: Dashboard: highlight performance scores in column A with three‑state traffic‑light icons while still showing the underlying figures. | Report conversion: take an existing XLSX report, add visual icon cues, and deliver it as an XLSM so downstream users can run embedded macros. | Automated KPI sheet: generate a weekly spreadsheet where icon sets flag critical values and the file must remain macro‑ready for further processing.
// AI Prompts: Generate C# code with Aspose.Cells that adds a TrafficLights31 icon set to A1:A10, shows the cell values, and saves the file as an XLSM with macros enabled. | Explain step‑by‑step how to turn on macro support in an Aspose.Cells workbook and configure an icon‑set conditional format that displays values. | Provide a concise tutorial for applying an icon set conditional formatting rule to a range and exporting the workbook as SaveFormat.Xlsm.

using System;
using Aspose.Cells;

// Loads an XLSX file, activates macro support, creates a conditional formatting rule that applies the TrafficLights31 icon set to cells A1:A10 while keeping the numeric values visible, and saves the result as a macro‑enabled XLSM document.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Enable macros so the workbook can be saved as XLSM
        workbook.Settings.EnableMacros = true;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a new conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection formatConditions = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the icon set will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 9,
            EndColumn = 0
        };
        formatConditions.AddArea(area);

        // Add an icon set condition
        int conditionIndex = formatConditions.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = formatConditions[conditionIndex];

        // Configure the icon set (e.g., TrafficLights31) and ensure cell values are shown
        condition.IconSet.Type = IconSetType.TrafficLights31;
        condition.IconSet.ShowValue = true;

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
