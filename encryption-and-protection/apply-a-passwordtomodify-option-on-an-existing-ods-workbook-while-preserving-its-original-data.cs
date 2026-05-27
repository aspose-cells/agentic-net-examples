using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    class ApplyWriteProtectionToOds
    {
        static void Main()
        {
            try
            {
                // Path to the existing ODS workbook
                string inputPath = "ExistingWorkbook.ods";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one sheet
                }

                // Set the password required to modify the file
                workbook.Settings.WriteProtection.Password = "ModifyPassword123";

                // Save the workbook with write‑protection applied
                string outputPath = "ProtectedWorkbook.ods";
                workbook.Save(outputPath, SaveFormat.Ods);

                Console.WriteLine("Workbook saved with write‑protection password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}