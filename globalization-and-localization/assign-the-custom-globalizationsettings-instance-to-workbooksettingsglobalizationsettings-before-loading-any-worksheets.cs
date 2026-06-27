using System;
using Aspose.Cells;
using System.Globalization;

namespace AsposeCellsCustomGlobalizationDemo
{
    // Custom globalization settings – override methods as needed
    public class MyGlobalizationSettings : GlobalizationSettings
    {
        // Example: change boolean display strings
        public override string GetBooleanValueString(bool value)
        {
            return value ? "YES_CUSTOM" : "NO_CUSTOM";
        }

        // Example: change error value strings
        public override string GetErrorValueString(string err)
        {
            // Map a few common errors to custom text
            return err switch
            {
                "#DIV/0!" => "#DIV/0!_CUSTOM",
                "#N/A" => "#N/A_CUSTOM",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (no worksheets are accessed yet)
            Workbook wb = new Workbook();

            // 2. Assign the custom globalization settings BEFORE any worksheet operations
            wb.Settings.GlobalizationSettings = new MyGlobalizationSettings();

            // 3. Now it is safe to work with worksheets – the settings will affect them
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data to demonstrate the custom settings
            cells[0, 0].PutValue(true);                     // Boolean – will use custom string
            cells[0, 1].PutValue(false);                    // Boolean – will use custom string
            cells[0, 2].PutValue("#DIV/0!");                // Error – custom text
            cells[0, 3].PutValue("#N/A");                   // Error – custom text
            cells[0, 4].PutValue(1234.56);                  // Numeric – unaffected

            // 4. Save the workbook to verify the result
            wb.Save("CustomGlobalizationDemo.xlsx");
        }
    }
}