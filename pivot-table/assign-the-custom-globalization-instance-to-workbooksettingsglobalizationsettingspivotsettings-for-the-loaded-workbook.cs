using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure a GlobalizationSettings instance exists
        if (workbook.Settings.GlobalizationSettings == null)
        {
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
        }

        // Assign custom pivot globalization settings
        workbook.Settings.GlobalizationSettings.PivotSettings = new CustomPivotGlobalizationSettings();

        // Save the workbook with the new settings
        workbook.Save("output.xlsx");
    }

    // Custom implementation of PivotGlobalizationSettings
    class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string GetTextOfTotal()
        {
            return "Custom Total";
        }

        public override string GetTextOfGrandTotal()
        {
            return "Custom Grand Total";
        }

        public override string GetTextOfDataFieldHeader()
        {
            return "Custom Data Header";
        }

        public override string GetTextOfProtectedName(string protectedName)
        {
            return protectedName + "_Custom";
        }
    }
}