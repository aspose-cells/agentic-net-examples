// Title: Assign Custom PivotGlobalizationSettings to a Workbook in Aspose.Cells (.NET)
// Description: Shows how to load an Excel file, ensure it has a GlobalizationSettings object, attach a user‑defined class derived from PivotGlobalizationSettings (overriding Total, Grand Total, Data Header and protected‑name texts), refresh every pivot table, and save the result using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | PivotGlobalizationSettings | custom pivot labels | Workbook GlobalizationSettings | PivotTable refresh | override total text | grand total customization | data field header text | protected name suffix
// Common Searches: Aspose.Cells assign custom PivotGlobalizationSettings | Change pivot total label with Aspose.Cells .NET | Set PivotSettings globally in a workbook | Refresh pivot tables after globalization change Aspose | C# example custom pivot globalization
// Developer Intent: Attach a user‑defined PivotGlobalizationSettings object to Workbook.Settings.GlobalizationSettings.PivotSettings and update all pivot tables.
// Use Cases: Standardize pivot terminology across a corporate workbook before distribution. | Localize pivot table captions for non‑English audiences with a single class. | Add a suffix to protected field names in every pivot table of an automated report. | Programmatically modify pivot UI text without manual Excel editing.
// AI Prompts: Generate C# code that defines a class inheriting from PivotGlobalizationSettings to customize Total, Grand Total, and Data Header texts, assigns it to Workbook.Settings.GlobalizationSettings.PivotSettings, refreshes each pivot table, and saves the workbook. | Explain why pivot tables must be refreshed after changing PivotGlobalizationSettings in Aspose.Cells. | Provide troubleshooting steps when custom pivot globalization texts do not appear in the saved Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Custom globalization settings for pivot tables
    // Shows how to load an Excel file, ensure it has a GlobalizationSettings object, attach a user‑defined class derived from PivotGlobalizationSettings (overriding Total, Grand Total, Data Header and protected‑name texts), refresh every pivot table, and save the result using Aspose.Cells for C#.
    public class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        // Override the text displayed for the "Total" label
        public override string GetTextOfTotal()
        {
            return "Custom Total";
        }

        // Override the text displayed for the "Grand Total" label
        public override string GetTextOfGrandTotal()
        {
            return "Custom Grand Total";
        }

        // Override the text displayed for the data field header
        public override string GetTextOfDataFieldHeader()
        {
            return "Custom Data Header";
        }

        // Override the text displayed for protected names (example)
        public override string GetTextOfProtectedName(string protectedName)
        {
            return protectedName + "_Custom";
        }
    }

    public class AssignPivotGlobalizationDemo
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure the workbook has a GlobalizationSettings instance
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();

            // Assign the custom pivot globalization settings
            workbook.Settings.GlobalizationSettings.PivotSettings = new CustomPivotGlobalizationSettings();

            // If the workbook contains pivot tables, refresh them to apply the new settings
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    pt.RefreshData();
                    pt.CalculateData();
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            AssignPivotGlobalizationDemo.Run();
        }
    }
}
