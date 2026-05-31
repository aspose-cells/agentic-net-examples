using System;
using Aspose.Cells;

namespace AsposeCellsCustomErrorHandler
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom globalization settings to replace #N/A with a user‑defined placeholder
            workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings("#N/A", "Data not available");

            // Create a formula that will produce #N/A (lookup value not found)
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=VLOOKUP(\"Missing\",B1:C2,2,FALSE)";

            // Populate the lookup table without the key "Missing"
            sheet.Cells["B1"].PutValue("Key1");
            sheet.Cells["C1"].PutValue(100);
            sheet.Cells["B2"].PutValue("Key2");
            sheet.Cells["C2"].PutValue(200);

            // Calculate formulas
            workbook.CalculateFormula();

            // Display the result; the custom error handler replaces #N/A
            Console.WriteLine("Cell A1 display value: " + cell.DisplayStringValue);

            // Save the workbook
            workbook.Save("CustomErrorHandlerDemo.xlsx");
        }
    }

    // Custom globalization settings that overrides error string mapping
    public class CustomErrorGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _targetError;
        private readonly string _replacement;

        public CustomErrorGlobalizationSettings(string targetError, string replacement)
        {
            _targetError = targetError;
            _replacement = replacement;
        }

        public override string GetErrorValueString(string err)
        {
            // Replace the specified error with the custom placeholder
            if (err == _targetError)
                return _replacement;

            // Fallback to default behavior for other errors
            return base.GetErrorValueString(err);
        }
    }
}