// Title: Reverting Aspose.Cells Workbook CultureInfo to Original Settings When Custom Label Code Throws an Exception (C#)
// Common Searches: how to catch CultureNotFoundException and reset workbook culture in Aspose.Cells C# | fallback to original CultureInfo after custom label failure in Aspose.Cells | restore default globalization settings for Aspose.Cells workbook when custom label throws | exception handling for workbook.Settings.CultureInfo in Aspose.Cells .NET | apply custom label settings with safe revert to original culture in Aspose.Cells
// Tags: Aspose.Cells workbook culture fallback | Aspose.Cells revert workbook culture | Aspose.Cells custom label exception handling | Aspose.Cells globalization error recovery | Aspose.Cells .NET cultureinfo reset

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a Workbook, saves its initial Settings.CultureInfo, attempts to apply custom label settings that may raise an exception (e.g., assigning an invalid culture), catches any error, restores the original CultureInfo, and then saves the workbook while logging success or failure.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Preserve the original globalization settings
        CultureInfo originalCulture = workbook.Settings.CultureInfo;

        try
        {
            // Attempt to apply custom label settings that may throw an exception
            ApplyCustomLabelSettings(workbook);
        }
        catch (Exception ex)
        {
            // If an exception occurs, revert to the original globalization settings
            workbook.Settings.CultureInfo = originalCulture;

            Console.WriteLine("Custom label settings failed: " + ex.Message);
            Console.WriteLine("Reverted to default globalization settings.");
        }

        try
        {
            // Save the workbook (lifecycle rule)
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save workbook: " + ex.Message);
        }
    }

    // Example custom label method that could throw an exception
    static void ApplyCustomLabelSettings(Workbook wb)
    {
        // Simulate an error by assigning an invalid culture name
        // This will throw a CultureNotFoundException
        wb.Settings.CultureInfo = new CultureInfo("invalid-culture");
    }
}
