// Title: Assign a custom PivotGlobalizationSettings instance to a workbook’s GlobalizationSettings.PivotSettings using Aspose.Cells for .NET
// AI Prompts: Create a class derived from PivotGlobalizationSettings that overrides GetTextOfTotal, GetTextOfGrandTotal, and GetTextOfDataFieldHeader, then assign an instance of this class to workbook.Settings.GlobalizationSettings.PivotSettings and refresh all pivot tables. | Load an existing XLSX file with Aspose.Cells, apply a custom pivot globalization object to the workbook’s settings, recalculate each pivot table, and save the modified workbook to a new file. | Write a script that scans a directory for Excel files, sets the same custom PivotGlobalizationSettings on each workbook’s globalization settings, refreshes their pivot tables, and writes the updated files back.
// Common Searches: how to apply custom pivot globalization settings to a workbook in Aspose.Cells C# | setting PivotSettings property on GlobalizationSettings for Excel pivot tables | refreshing pivot tables after changing globalization labels Aspose.Cells | override total and grand total text in Excel pivot tables programmatically | apply custom pivot table text to multiple workbooks using Aspose.Cells
// Tags: pivot globalization settings customization Aspose.Cells | set workbook globalization pivot settings C# | override pivot total label Aspose.Cells | recalculate pivot tables after globalization update | apply custom pivot settings to batch of workbooks | PivotGlobalizationSettings subclass example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsCustomPivotGlobalization
{
    // Custom globalization settings for pivot tables
    // Shows how to subclass PivotGlobalizationSettings to customize pivot table labels, assign the custom instance to a workbook's GlobalizationSettings.PivotSettings, refresh all pivot tables, and save the updated workbook.
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
            return protectedName + "_Protected";
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Create a new GlobalizationSettings instance if not already set
            GlobalizationSettings globalization = new GlobalizationSettings();

            // Assign the custom pivot globalization settings
            globalization.PivotSettings = new CustomPivotGlobalizationSettings();

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Refresh all pivot tables in the workbook to apply the new settings
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    pivot.RefreshData();
                    pivot.CalculateData();
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
