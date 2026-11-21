// Title: How to fallback to invariant CultureInfo in Aspose.Cells when setting a custom culture throws an exception (C#)
// Common Searches: C# Aspose.Cells set workbook culture with try-catch fallback to invariant | handle exception when assigning workbook.Settings.CultureInfo in Aspose.Cells | default globalization settings for Aspose.Cells workbook if custom culture fails | example of using CultureInfo.InvariantCulture as fallback in Aspose.Cells | Aspose.Cells revert to default culture after CultureInfo error in .NET
// Tags: Aspose.Cells cultureinfo fallback | Workbook.Settings.CultureInfo exception handling | C# invariant culture fallback Aspose.Cells | custom globalization settings Aspose.Cells | Excel export default culture Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a Workbook, writes a value, attempts to assign a custom French CultureInfo to workbook.Settings.CultureInfo, and uses a try‑catch block to revert to CultureInfo.InvariantCulture if any exception occurs, then saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and set a sample value
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Attempt to apply custom globalization settings
            try
            {
                // Set a custom culture (e.g., French - France)
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                // Note: Custom boolean string properties are not available in the current Aspose.Cells version.
            }
            catch (Exception)
            {
                // Fallback to default globalization settings if any custom method fails
                workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;
            }

            // Save the workbook (lifecycle rule)
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
