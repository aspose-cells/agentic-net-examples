using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FreezeAndProtectHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the first row (header) so it stays visible while scrolling.
                // "A2" means the freeze line is just below row 1; freeze 1 row, 0 columns.
                sheet.FreezePanes("A2", 1, 0);

                // Protect the worksheet with a password.
                // The third argument (oldPassword) is required in newer API versions; pass null when not changing it.
                sheet.Protect(ProtectionType.All, "HeaderPassword123", null);

                // Define output path and ensure the directory exists
                string outputPath = "FreezeAndProtectedHeader.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FreezeAndProtectHeader.Run();
        }
    }
}