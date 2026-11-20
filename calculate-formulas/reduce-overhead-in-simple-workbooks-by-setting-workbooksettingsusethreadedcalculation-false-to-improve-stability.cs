// Title: Disable threaded calculation in Aspose.Cells (C#) to lower overhead for simple workbooks
// Common Searches: C# Aspose.Cells how to turn off threaded calculation for faster workbook save | setting Workbook.Settings.UseThreadedCalculation to false in .NET Excel processing | improve performance of small Excel files by disabling calculation threading in Aspose.Cells | Aspose.Cells calculation engine thread usage control in C# applications | troubleshoot high CPU usage in Aspose.Cells by disabling threaded calculation
// Tags: disable threaded calculation Aspose.Cells | Workbook.Settings.UseThreadedCalculation false C# | Aspose.Cells calculation engine performance tuning | simple workbook overhead reduction Aspose.Cells | C# Excel processing thread control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to set Workbook.Settings.UseThreadedCalculation to false in a C# Aspose.Cells workbook, which can reduce CPU overhead and improve stability when working with simple Excel files. It also includes directory preparation and error handling before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                var workbook = new Workbook();

                // Threaded calculation disabling is not required / not available in this version
                // workbook.Settings.UseThreadedCalculation = false;

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
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
}
