// Title: C# – Aspose.Cells custom GlobalizationSettings to replace #N/A with a user‑defined placeholder
// Description: Demonstrates how to subclass Aspose.Cells.GlobalizationSettings, override GetErrorValueString to return a developer‑specified string for the "#N/A" error, apply the settings to a workbook, insert an NA() formula, calculate it, and save the file with the custom placeholder displayed.
// Keywords: Aspose.Cells | C# | .NET | custom error handling | #N/A placeholder | GlobalizationSettings | GetErrorValueString | override error display | Excel error customization | NA() function
// Common Searches: Aspose.Cells replace #N/A error | custom GlobalizationSettings Aspose.Cells .NET | override GetErrorValueString example | display custom text for NA() error | Aspose.Cells error placeholder C# | hide #N/A in generated Excel workbook
// Developer Intent: Show a custom message instead of the default #N/A error when cells are displayed.
// Use Cases: Present user‑friendly messages such as "Data Not Available" in reports. | Localize error messages by supplying language‑specific placeholders per workbook. | Suppress calculation errors in exported Excel files while keeping the original formulas intact.
// AI Prompts: Write a C# class that inherits from Aspose.Cells.GlobalizationSettings to replace #N/A errors with a custom string and show how to use it. | Generate sample code that applies a custom GlobalizationSettings to a workbook, adds an NA() formula, calculates it, and saves the file with the placeholder text. | Explain how to extend the custom error handler to also replace other errors like #DIV/0! or #VALUE! in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomErrorHandler
{
    // Custom globalization settings to replace #N/A with a user‑defined placeholder
    // Demonstrates how to subclass Aspose.Cells.GlobalizationSettings, override GetErrorValueString to return a developer‑specified string for the "#N/A" error, apply the settings to a workbook, insert an NA() formula, calculate it, and save the file with the custom placeholder displayed.
    public class CustomErrorGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _naPlaceholder;

        public CustomErrorGlobalizationSettings(string naPlaceholder)
        {
            _naPlaceholder = naPlaceholder;
        }

        public override string GetErrorValueString(string err)
        {
            // Replace only the #N/A error; other errors fall back to default handling
            return err switch
            {
                "#N/A" => _naPlaceholder,
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the placeholder text for #N/A errors
            string placeholder = "Data Not Available";

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings(placeholder);

            // Create a cell that will contain the #N/A error using the NA() function
            Cell errorCell = sheet.Cells["A1"];
            errorCell.Formula = "=NA()";

            // Calculate formulas so the error is generated
            workbook.CalculateFormula();

            // Retrieve the display string; it will use the custom placeholder
            string displayedValue = errorCell.DisplayStringValue;

            Console.WriteLine($"Cell A1 displayed value: {displayedValue}");

            // Save the workbook (lifecycle save)
            workbook.Save("CustomErrorHandlerDemo.xlsx");
        }
    }
}
