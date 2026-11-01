// Title: Disable the Precision As Displayed setting in an Aspose.Cells workbook to enable full double‑precision calculations (C#)
// Common Searches: Aspose.Cells C# disable Precision As Displayed workbook setting | set workbook calculation precision to full double in Aspose.Cells .NET | turn off Precision As Displayed mode using Aspose.Cells WorkbookSettings | use reflection to change IsPrecisionAsDisplayed property in Aspose.Cells | example of full precision calculations with Aspose.Cells C#
// Tags: disable workbook precision mode Aspose.Cells | full double precision calculations .NET | WorkbookSettings precision property reflection | Aspose.Cells settings via reflection | C# calculation precision control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, uses reflection to locate the optional IsPrecisionAsDisplayed property on WorkbookSettings, sets it to false, and saves the file. This disables the Precision As Displayed mode so that all formula evaluations use full double‑precision arithmetic.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Attempt to disable "Precision as displayed" if the property exists in this version
            var settings = workbook.Settings;
            var prop = typeof(WorkbookSettings).GetProperty("IsPrecisionAsDisplayed");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(settings, false);
            }

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
