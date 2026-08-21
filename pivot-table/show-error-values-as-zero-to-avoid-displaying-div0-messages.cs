// Title: Show #DIV/0! as Zero in C# with AspNet.Cells GlobalizationSettings
// Description: Provides a C# example that defines a custom GlobalizationSettings class for Aspose.Cells, overrides GetErrorValueString to convert the Excel #DIV/0! error to "0", applies the setting to a workbook, recalculates formulas, and saves the file.
// Keywords: Aspose.Cells C# | GlobalizationSettings | replace #DIV/0! with zero | Excel error handling | custom error display | formula calculation | workbook save | Aspose.Cells example | division by zero error | override GetErrorValueString
// Common Searches: Aspose.Cells replace #DIV/0! with 0 | C# GlobalizationSettings error value string | show zero instead of division error Aspose.Cells | custom error handling in Aspose.Cells workbook | override GetErrorValueString Aspose.Cells C#
// Developer Intent: Use Aspose.Cells to display zero instead of the #DIV/0! error in calculated cells.
// Use Cases: Generate reports where division‑by‑zero cells should appear as 0 for readability. | Apply a single GlobalizationSettings instance to enforce consistent error display across all worksheets. | Extend the overridden method to map additional Excel errors (e.g., #N/A, #VALUE!) to user‑defined values.
// AI Prompts: Write C# code that uses Aspose.Cells to replace all Excel error values with custom strings via GlobalizationSettings. | Show how to map #N/A and #VALUE! to empty strings while keeping #DIV/0! as zero in Aspose.Cells. | Explain how to attach a custom GlobalizationSettings object to an existing workbook without triggering a full recalculation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings to replace division‑by‑zero errors with zero
    // Provides a C# example that defines a custom GlobalizationSettings class for Aspose.Cells, overrides GetErrorValueString to convert the Excel #DIV/0! error to "0", applies the setting to a workbook, recalculates formulas, and saves the file.
    public class ZeroErrorGlobalizationSettings : GlobalizationSettings
    {
        public override string GetErrorValueString(string err)
        {
            // Map the specific error string to "0"; otherwise use the default handling
            return err == "#DIV/0!" ? "0" : base.GetErrorValueString(err);
        }
    }

    public class ShowErrorAsZeroDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply the custom globalization settings
            workbook.Settings.GlobalizationSettings = new ZeroErrorGlobalizationSettings();

            // Create a cell with a formula that would normally produce #DIV/0!
            Cell errorCell = sheet.Cells["A1"];
            errorCell.Formula = "=1/0";

            // Calculate formulas (the error will be transformed to "0")
            workbook.CalculateFormula();

            // Verify the displayed value
            Console.WriteLine("Cell A1 display value: " + errorCell.DisplayStringValue); // Expected output: 0

            // Save the workbook (lifecycle rule: save)
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
