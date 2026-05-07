using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAdvancedDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Read a large workbook (standard loading for simplicity)
            // -----------------------------------------------------------------
            string largeFilePath = "LargeFile.xlsx";
            CreateSampleLargeWorkbook(largeFilePath, 5000, 5); // create a sample file

            Workbook largeWorkbook = new Workbook(largeFilePath);
            Console.WriteLine("[Info] Finished reading of large workbook.");

            // -----------------------------------------------------------------
            // 2. Merge multiple XLS files into a single workbook
            // -----------------------------------------------------------------
            string[] filesToMerge = { "File1.xls", "File2.xls" };
            CreateSampleXlsFile(filesToMerge[0], "First file content");
            CreateSampleXlsFile(filesToMerge[1], "Second file content");

            string cacheFile = "MergeCache.tmp";
            string mergedFile = "MergedOutput.xls";

            // Merge using CellsHelper (data, style, formulas are merged)
            CellsHelper.MergeFiles(filesToMerge, cacheFile, mergedFile);
            Console.WriteLine("[Info] Files merged into: " + mergedFile);

            // -----------------------------------------------------------------
            // 3. Convert merged XLS to PDF using ConversionUtility
            // -----------------------------------------------------------------
            string pdfFile = "MergedOutput.pdf";
            ConversionUtility.Convert(mergedFile, pdfFile);
            Console.WriteLine("[Info] Converted merged XLS to PDF: " + pdfFile);

            // -----------------------------------------------------------------
            // 4. Save workbook with XlsSaveOptions (performance‑oriented settings)
            // -----------------------------------------------------------------
            Workbook optWorkbook = new Workbook();
            Worksheet optSheet = optWorkbook.Worksheets[0];
            for (int i = 0; i < 100; i++)
                optSheet.Cells[i, 0].PutValue($"Item {i}");

            XlsSaveOptions saveOptions = new XlsSaveOptions
            {
                MatchColor = true,
                ValidateMergedAreas = true,
                RefreshChartCache = true,
                ClearData = false,
                CreateDirectory = true
            };
            string optFile = "AdvancedOptions.xls";
            optWorkbook.Save(optFile, saveOptions);
            Console.WriteLine("[Info] Workbook saved with XlsSaveOptions: " + optFile);

            // -----------------------------------------------------------------
            // 5. Save a very large workbook (standard save for simplicity)
            // -----------------------------------------------------------------
            Workbook lightSaveWorkbook = new Workbook();
            string lightSavedFile = "LightCellsSaved.xls";
            lightSaveWorkbook.Save(lightSavedFile, SaveFormat.Excel97To2003);
            Console.WriteLine("[Info] Large workbook saved: " + lightSavedFile);

            // -----------------------------------------------------------------
            // Cleanup temporary files (optional)
            // -----------------------------------------------------------------
            TryDeleteFile(largeFilePath);
            TryDeleteFile(filesToMerge[0]);
            TryDeleteFile(filesToMerge[1]);
            TryDeleteFile(cacheFile);
            // Note: mergedFile, pdfFile, optFile, lightSavedFile are kept for inspection
        }

        private static void CreateSampleLargeWorkbook(string path, int rows, int cols)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    ws.Cells[r, c].PutValue($"R{r}C{c}");
            wb.Save(path, SaveFormat.Xlsx);
        }

        private static void CreateSampleXlsFile(string path, string content)
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue(content);
            wb.Save(path, SaveFormat.Excel97To2003);
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // ignore any errors during cleanup
            }
        }
    }
}