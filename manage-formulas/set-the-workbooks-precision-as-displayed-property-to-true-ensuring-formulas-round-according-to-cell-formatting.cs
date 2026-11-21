// Title: How to enable Precision As Displayed for a workbook using Aspose.Cells in C#
// Common Searches: Aspose.Cells set precision as displayed property in C# | Enable workbook precision as displayed to round formulas in Aspose.Cells .NET | Use Aspose.Cells Settings to force formula rounding according to cell format | C# code to turn on precision as displayed in an Aspose.Cells workbook | Reflection workaround for IsPrecisionAsDisplayed property in older Aspose.Cells versions
// Tags: Aspose.Cells workbook precision as displayed | C# set IsPrecisionAsDisplayed via reflection | formula rounding based on cell formatting Aspose.Cells | Aspose.Cells Settings precision option .NET | enable precision as displayed programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, uses reflection to locate the Settings.IsPrecisionAsDisplayed property (when available), sets it to true to enable Precision As Displayed, and saves the file, ensuring formulas are rounded according to cell formatting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Attempt to enable "Precision as displayed" via reflection
            // (property may not exist in some library versions)
            var settings = workbook.Settings;
            var precisionProp = settings.GetType().GetProperty("IsPrecisionAsDisplayed");
            if (precisionProp != null && precisionProp.CanWrite)
            {
                precisionProp.SetValue(settings, true);
            }

            // Define output file path
            string outputPath = "PrecisionAsDisplayed.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
