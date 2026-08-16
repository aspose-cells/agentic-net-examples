// Title: C# Aspose.Cells: Custom error handler to replace #N/A with a user‑defined placeholder
// Description: Shows how to create a subclass of GlobalizationSettings in Aspose.Cells for .NET, override GetErrorValueString, and apply the custom settings to a workbook so that any #N/A error (for example from a VLOOKUP miss) is shown as a developer‑specified text such as "Data Not Available". The sample calculates formulas, reads the display string, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | CustomErrorGlobalizationSettings | GlobalizationSettings | GetErrorValueString | #N/A replacement | custom placeholder | error handling | VLOOKUP error | Workbook.Settings.GlobalizationSettings | display string | Excel export
// Common Searches: Aspose.Cells replace #N/A with custom text | override GetErrorValueString Aspose.Cells | custom globalization settings error value | C# Aspose.Cells custom error handler example | how to change #N/A display in Aspose.Cells
// Developer Intent: Replace #N/A error values with a custom placeholder when displaying cell contents.
// Use Cases: Show a friendly message like "Data Not Available" for missing lookup results. | Provide localized error messages across an entire workbook by defining language‑specific placeholders. | Ensure consistent placeholder text for all #N/A errors before exporting the workbook to Excel.
// AI Prompts: Write a GlobalizationSettings subclass that substitutes "#N/A" with "Not Found" and apply it to an Aspose.Cells workbook. | Explain step‑by‑step how to assign a custom error handler to Workbook.Settings.GlobalizationSettings and retrieve the display string of a cell containing an error. | Generate sample C# code that uses a custom placeholder for #N/A errors in a workbook with multiple formulas and saves the result.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomErrorHandler
{
    // Custom globalization settings to replace #N/A with a user‑defined placeholder
    // Shows how to create a subclass of GlobalizationSettings in Aspose.Cells for .NET, override GetErrorValueString, and apply the custom settings to a workbook so that any #N/A error (for example from a VLOOKUP miss) is shown as a developer‑specified text such as "Data Not Available". The sample calculates formulas, reads the display string, and saves the workbook.
    public class CustomErrorGlobalizationSettings : GlobalizationSettings
    {
        // Placeholder text to display instead of #N/A
        private readonly string _naPlaceholder;

        public CustomErrorGlobalizationSettings(string naPlaceholder)
        {
            _naPlaceholder = naPlaceholder;
        }

        // Override the method that returns the display string for an error value
        public override string GetErrorValueString(string err)
        {
            // If the error is #N/A, return the custom placeholder; otherwise use default behavior
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Apply the custom globalization settings (replace #N/A with "Data Not Available")
            workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings("Data Not Available");

            // Example that generates #N/A: VLOOKUP with a missing key
            cells["A1"].PutValue("Key");
            cells["A2"].PutValue("Value1");
            cells["B1"].Formula = "=VLOOKUP(\"MissingKey\",A1:A2,1,FALSE)";

            // Calculate formulas so that the error is produced
            workbook.CalculateFormula();

            // Retrieve the display string; it will use the custom placeholder
            string result = cells["B1"].DisplayStringValue;
            Console.WriteLine($"Cell B1 display value: {result}");

            // Save the workbook
            workbook.Save("CustomErrorHandlerDemo.xlsx");
        }
    }
}
