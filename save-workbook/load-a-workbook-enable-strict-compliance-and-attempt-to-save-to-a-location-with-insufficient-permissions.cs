// Title: Load an Excel workbook, enable ISO/IEC 29500:2008 strict compliance, and handle save errors caused by insufficient permissions using Aspose.Cells in C#
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, sets workbook.Settings.Compliance to OoxmlCompliance.Iso29500_2008_Strict, and attempts to save it to a directory without write rights, capturing any exception. | Show how to catch and log an UnauthorizedAccessException when saving a strict‑compliance workbook to a protected system folder in C# with Aspose.Cells.
// Common Searches: Aspose.Cells C# enable ISO 29500 strict mode and save to protected folder | how to catch unauthorized access exception when saving Excel file with Aspose.Cells | save workbook with strict compliance to System32 directory C# | set OoxmlCompliance.Iso29500_2008_Strict before saving and handle permission error
// Tags: enable strict OoxmlCompliance Aspose.Cells | save workbook to protected directory C# | handle UnauthorizedAccessException Aspose.Cells | load workbook then set ISO29500 strict mode | Aspose.Cells permission error on save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStrictComplianceDemo
{
    // The example creates a temporary workbook, reloads it, switches the workbook's compliance to ISO/IEC 29500:2008 Strict, and then tries to save the file to a system folder that requires elevated rights. It demonstrates catching the resulting permission‑related exception and cleaning up the temporary file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook tempWorkbook = new Workbook();
            tempWorkbook.Worksheets[0].Cells["A1"].PutValue("Strict compliance test");

            // Save the workbook to a temporary file (normal location)
            string tempPath = Path.Combine(Path.GetTempPath(), "TempWorkbook.xlsx");
            tempWorkbook.Save(tempPath);
            tempWorkbook.Dispose();

            // Load the workbook from the temporary file
            Workbook workbook = new Workbook(tempPath);

            // Enable ISO/IEC 29500:2008 Strict compliance
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Attempt to save to a location with insufficient permissions
            // Example: system directory (usually requires elevated rights)
            string restrictedPath = @"C:\Windows\System32\RestrictedWorkbook.xlsx";

            try
            {
                workbook.Save(restrictedPath);
                Console.WriteLine("Workbook saved successfully (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook due to insufficient permissions:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                workbook.Dispose();

                // Clean up temporary file
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }
    }
}
