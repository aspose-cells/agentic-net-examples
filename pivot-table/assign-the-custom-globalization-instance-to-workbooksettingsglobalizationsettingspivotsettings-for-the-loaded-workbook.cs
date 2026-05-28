using System;
using Aspose.Cells;
using Aspose.Cells.Settings;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure a GlobalizationSettings instance exists
        if (workbook.Settings.GlobalizationSettings == null)
        {
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
        }

        // Create a custom PivotGlobalizationSettings instance
        var customPivotSettings = new CustomPivotGlobalizationSettings();

        // Assign the custom instance to the PivotSettings property
        workbook.Settings.GlobalizationSettings.PivotSettings = customPivotSettings;

        // Refresh all pivot tables so they pick up the new globalization settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pt in sheet.PivotTables)
            {
                pt.RefreshData();
                pt.CalculateData();
            }
        }

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }

    // Custom implementation of PivotGlobalizationSettings
    class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string GetTextOfTotal()
        {
            return "My Total";
        }

        public override string GetTextOfGrandTotal()
        {
            return "My Grand Total";
        }

        public override string GetTextOfDataFieldHeader()
        {
            return "My Data Header";
        }

        public override string GetTextOfProtectedName(string protectedName)
        {
            return protectedName + "_custom";
        }
    }
}