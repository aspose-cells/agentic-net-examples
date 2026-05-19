using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLicenseDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Apply the Aspose.Cells license if the file exists.
                const string licensePath = "Aspose.Cells.NET.lic";
                if (File.Exists(licensePath))
                {
                    try
                    {
                        var license = new License();
                        license.SetLicense(licensePath);
                        Console.WriteLine("License applied successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to apply license: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"License file not found at '{licensePath}'. Running in evaluation mode.");
                }

                // Verify that the license (if any) has been applied.
                Console.WriteLine($"Workbook licensed: {new Workbook().IsLicensed}");

                // Create a new workbook and add sample data.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Licensed workbook - no watermarks");

                // Save the workbook.
                const string outputPath = "LicensedWorkbook.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}