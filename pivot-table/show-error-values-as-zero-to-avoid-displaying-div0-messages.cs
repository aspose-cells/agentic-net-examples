using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings to replace error strings with "0"
    public class ZeroErrorGlobalizationSettings : GlobalizationSettings
    {
        public override string GetErrorValueString(string err)
        {
            // Map division by zero error to zero, keep other errors unchanged
            return err == "#DIV/0!" ? "0" : base.GetErrorValueString(err);
        }
    }

    public class ShowErrorAsZeroDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply custom globalization settings
            workbook.Settings.GlobalizationSettings = new ZeroErrorGlobalizationSettings();

            // Cell with a formula that causes division by zero error
            Cell errorCell = sheet.Cells["A1"];
            errorCell.Formula = "=1/0";

            // Calculate formulas (errors will be generated)
            workbook.CalculateFormula();

            // Retrieve the display string; thanks to the custom settings it will be "0"
            string displayedValue = errorCell.DisplayStringValue;
            Console.WriteLine($"Cell A1 displayed value: {displayedValue}");

            // Save the workbook
            workbook.Save("ShowErrorAsZeroDemo.xlsx");
        }
    }

    class Program
    {
        static void Main()
        {
            ShowErrorAsZeroDemo.Run();
        }
    }
}