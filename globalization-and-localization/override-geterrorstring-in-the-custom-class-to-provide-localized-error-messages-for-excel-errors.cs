using System;
using Aspose.Cells;

// Custom globalization settings that provide localized error messages
public class CustomErrorGlobalizationSettings : GlobalizationSettings
{
    // Override the method that returns the display string for an error value
    public override string GetErrorValueString(string err)
    {
        // Map specific Excel error strings to custom, localized messages
        return err switch
        {
            "#DIV/0!" => "Custom Division Error",
            "#VALUE!" => "Custom Type Mismatch",
            "#NAME?" => "Custom Identifier Error",
            "#N/A"   => "Custom Not Available",
            _        => base.GetErrorValueString(err) // fallback to default behavior
        };
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook (lifecycle start)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Assign the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings();

        // Create a cell that will generate an error (division by zero)
        Cell errorCell = sheet.Cells["A1"];
        errorCell.Formula = "=1/0"; // This will produce the #DIV/0! error

        // Calculate formulas so the error is evaluated
        workbook.CalculateFormula();

        // Retrieve the custom error string using the overridden method
        string customError = errorCell.DisplayStringValue; // or errorCell.StringValue

        // Output the custom error message to the console
        Console.WriteLine("Custom error display: " + customError);

        // Save the workbook (lifecycle end)
        workbook.Save("CustomErrorGlobalization.xlsx");
    }
}