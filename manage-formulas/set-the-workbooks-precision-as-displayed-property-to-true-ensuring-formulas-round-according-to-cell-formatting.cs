// Title: Enable "Precision as Displayed" in an Aspose.Cells Workbook Using C# Reflection and Save as XLSX
// AI Prompts: Write C# code that uses reflection to set Workbook.Settings.IsPrecisionAsDisplayed to true, with fallback handling when the property does not exist. | Create a new Aspose.Cells workbook, activate precision‑as‑displayed mode via reflection, and persist the result to an XLSX file. | Generate a C# example that checks for the IsPrecisionAsDisplayed setting in Aspose.Cells, toggles it on, and includes error handling for older library versions.
// Common Searches: C# Aspose.Cells how to turn on precision as displayed for formulas | using reflection to set IsPrecisionAsDisplayed property in Aspose.Cells workbook | enable formula rounding based on cell formatting in Aspose.Cells .NET | workaround for missing precision as displayed setting in older Aspose.Cells versions
// Tags: Aspose.Cells activate precision mode using reflection | C# dynamic workbook settings modification | formula rounding according to cell format Aspose.Cells | fallback for missing IsPrecisionAsDisplayed property | export workbook with precision rounding to XLSX

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, attempts to enable the Settings.IsPrecisionAsDisplayed flag via reflection (if the property exists in the current Aspose.Cells version), and saves the workbook as an XLSX file, while gracefully handling missing‑property scenarios and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Attempt to enable "Precision as displayed" via reflection (property may not exist in older versions)
                try
                {
                    PropertyInfo precisionProp = workbook.Settings.GetType().GetProperty("IsPrecisionAsDisplayed");
                    if (precisionProp != null && precisionProp.CanWrite)
                    {
                        precisionProp.SetValue(workbook.Settings, true);
                        Console.WriteLine("Precision as displayed enabled.");
                    }
                    else
                    {
                        Console.WriteLine("Precision as displayed setting is unavailable in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while setting precision as displayed: " + ex.Message);
                }

                // Define output path and ensure the directory exists
                string outputPath = "PrecisionWorkbook.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while saving the workbook: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
