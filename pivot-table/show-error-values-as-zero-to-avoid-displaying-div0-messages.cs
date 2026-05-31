using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that replace any error string with "0"
    public class ZeroErrorGlobalizationSettings : GlobalizationSettings
    {
        public override string GetErrorValueString(string err)
        {
            // Return "0" for any error (e.g., #DIV/0!, #VALUE!, etc.)
            return "0";
        }
    }

    public class ShowErrorAsZeroDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply custom globalization settings to replace errors with zero
            workbook.Settings.GlobalizationSettings = new ZeroErrorGlobalizationSettings();

            // Create a cell with a formula that generates a division by zero error
            Cell errorCell = sheet.Cells["A1"];
            errorCell.Formula = "=1/0";

            // Calculate formulas (errors will be converted to "0" by the custom settings)
            workbook.CalculateFormula();

            // Display the resulting value (should be "0")
            Console.WriteLine("Cell A1 display value after error handling: " + errorCell.DisplayStringValue);

            // Save the workbook (optional, demonstrates that the displayed value is zero)
            workbook.Save("ErrorAsZeroDemo.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ShowErrorAsZeroDemo.Run();
        }
    }
}