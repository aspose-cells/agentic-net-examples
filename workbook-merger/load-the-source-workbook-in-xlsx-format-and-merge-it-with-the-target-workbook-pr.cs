using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Determine the working directory
            string workDir = Directory.GetCurrentDirectory();

            // Paths to the source and target workbooks (XLSX format)
            string sourcePath = Path.Combine(workDir, "source.xlsx");
            string targetPath = Path.Combine(workDir, "target.xlsx");

            // Ensure source workbook exists
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "SourceSheet";
                wb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Ensure target workbook exists
            if (!File.Exists(targetPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "TargetSheet";
                wb.Save(targetPath, SaveFormat.Xlsx);
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Load the target workbook
            Workbook targetWorkbook = new Workbook(targetPath);

            // Merge the source workbook into the target workbook
            targetWorkbook.Combine(sourceWorkbook);

            // Save the merged workbook
            string mergedPath = Path.Combine(workDir, "merged.xlsx");
            targetWorkbook.Save(mergedPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Saved to '{mergedPath}'.");
        }
    }
}