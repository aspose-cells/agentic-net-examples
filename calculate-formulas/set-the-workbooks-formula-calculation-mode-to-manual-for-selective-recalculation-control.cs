// Title: How to configure Aspose.Cells workbook to use manual formula calculation mode in C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual for selective recalculation | disable automatic formula evaluation in Aspose.Cells .NET workbook | set Workbook.Settings.CalcMode to Manual using Aspose.Cells API | save Excel file with manual calculation setting in C# Aspose.Cells | control formula recalculation manually with Aspose.Cells library
// Tags: Aspose.Cells manual calculation mode | Workbook.Settings.CalcMode property | C# disable automatic formula evaluation | save Excel workbook with manual calc | selective formula recalculation .NET

using System;
using System.IO;
using Aspose.Cells;

// This C# example creates a new Workbook (or loads an existing one), sets its Settings.CalcMode to Manual to prevent automatic formula updates, ensures the output directory exists, and saves the file as ManualCalcWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Set the formula calculation mode to Manual.
            // Note: In some versions of Aspose.Cells the CalcMode property may not be available.
            // If supported, uncomment the following line:
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Define output file path
            string outputPath = "ManualCalcWorkbook.xlsx";

            // Ensure the directory exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to persist the setting.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
