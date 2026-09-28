// Title: How to enable the 1904 date system for legacy Excel formulas using Aspose.Cells in C#
// AI Prompts: Generate a C# example that switches an Aspose.Cells workbook to the 1904 date system before saving. | Provide C# code that uses reflection to set Settings.Is1904DateSystem = true on an Aspose.Cells workbook. | Show how to configure a new workbook with the 1904 date system enabled for legacy date‑dependent formulas in Aspose.Cells.
// Common Searches: C# Aspose.Cells set workbook to 1904 date system for old Excel files | How to configure 1904 date system in Aspose.Cells when using reflection | Enable legacy Excel date handling in a workbook with Aspose.Cells C# | Aspose.Cells 1904 date system example code for .NET | Set Is1904DateSystem property programmatically in Aspose.Cells C#
// Tags: Aspose.Cells Settings.Is1904DateSystem | C# reflection set workbook date system | legacy Excel date system Aspose.Cells | 1904 date system compatibility .NET | enable 1904 date system workbook

using System;
using Aspose.Cells;

// Creates a new Aspose.Cells workbook, uses reflection to set Settings.Is1904DateSystem to true for legacy date‑dependent formulas, and saves the file as "LegacyDateWorkbook.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Enable the 1904 date system for legacy date‑dependent formulas (using reflection for compatibility)
            try
            {
                var settings = workbook.Settings;
                var prop = settings.GetType().GetProperty("Is1904DateSystem");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settings, true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to set 1904 date system. {ex.Message}");
            }

            // Save the workbook to a file
            string outputPath = "LegacyDateWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
