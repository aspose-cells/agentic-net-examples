using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class ProtectMergedWorkbook
    {
        public static void Run()
        {
            try
            {
                const string targetPath = "Book1.xlsx";
                const string sourcePath = "Book2.xlsx";
                const string outputPath = "MergedProtected.xlsx";
                const string password = "MySecretPassword";

                // Verify input files exist
                if (!File.Exists(targetPath))
                    throw new FileNotFoundException($"Target workbook not found: {targetPath}");
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

                // Load workbooks
                using (var mergedWorkbook = new Workbook(targetPath))
                using (var sourceWorkbook = new Workbook(sourcePath))
                {
                    // Copy each worksheet from source to target
                    foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        // AddCopy expects the source sheet name
                        mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                    }

                    // Protect workbook structure
                    mergedWorkbook.Protect(ProtectionType.Structure, password);

                    // Save merged and protected workbook
                    mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                }

                Console.WriteLine($"Workbook merged and protected successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectMergedWorkbook.Run();
        }
    }
}