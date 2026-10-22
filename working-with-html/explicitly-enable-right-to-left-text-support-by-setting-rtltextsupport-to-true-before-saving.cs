// Title: How to activate right-to-left text support in an Aspose.Cells workbook before saving to XLSX with C#
// AI Prompts: Use reflection to turn on the RtlTextSupport flag in the Workbook.Settings object, then call Workbook.Save to generate the XLSX file. | Write a C# snippet that checks for the RtlTextSupport property, sets it to true, ensures the destination folder exists, and saves the workbook. | Demonstrate enabling bidirectional text layout in an Aspose.Cells workbook and exporting it as an Excel file.
// Common Searches: C# Aspose.Cells enable right-to-left layout before saving workbook | How to set RTL support in Aspose.Cells Excel export | Aspose.Cells workbook Settings RtlTextSupport property usage | Saving Excel file with bidirectional text using Aspose.Cells C#
// Tags: Aspose.Cells workbook Settings RtlTextSupport | enable RTL layout in XLSX via Aspose.Cells | C# reflection set Aspose.Cells property | ensure output directory before saving Excel file | right-to-left text support in generated Excel

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Aspose.Cells workbook, uses reflection to set the RtlTextSupport setting to true when available, guarantees the output folder exists, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable right-to-left text support if the property exists (handled via reflection)
            var settings = workbook.Settings;
            var rtlProp = settings.GetType().GetProperty("RtlTextSupport");
            if (rtlProp != null && rtlProp.CanWrite)
            {
                rtlProp.SetValue(settings, true);
            }

            // Determine output path and ensure its directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
