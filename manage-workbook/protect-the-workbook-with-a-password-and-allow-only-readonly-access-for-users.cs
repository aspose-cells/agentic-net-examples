using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectWorkbookReadOnlyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // (Optional) Add some data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Read‑only protected workbook");

                // Set write‑protection password and enable "Read‑only recommended"
                workbook.Settings.WriteProtection.Password = "readOnlyPwd";
                workbook.Settings.WriteProtection.RecommendReadOnly = true;

                // Define output file path
                string outputPath = "ReadOnlyProtectedWorkbook.xlsx";

                // If the file already exists, delete it to avoid IOException
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                // Save the workbook; users will be prompted to open as read‑only unless they provide the password
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorkbookReadOnlyDemo.Run();
        }
    }
}