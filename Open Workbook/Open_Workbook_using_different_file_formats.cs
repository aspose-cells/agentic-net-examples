using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOpenDifferentFormats
{
    class Program
    {
        static void Main()
        {
            // Paths to sample files in various formats.
            // Replace these with actual file locations when running the code.
            string xlsxPath = "sample.xlsx";
            string xlsPath = "sample.xls";
            string csvPath = "sample.csv";
            string jsonPath = "sample.json";
            string epubPath = "sample.epub";

            // -----------------------------------------------------------------
            // 1. Open an XLSX file using the simple string constructor.
            // -----------------------------------------------------------------
            Workbook workbookXlsx = new Workbook(xlsxPath);
            Console.WriteLine($"Opened XLSX workbook. Worksheets count: {workbookXlsx.Worksheets.Count}");
            workbookXlsx.Save("output_from_xlsx.pdf", SaveFormat.Pdf);
            Console.WriteLine("Saved XLSX workbook as PDF.");

            // -----------------------------------------------------------------
            // 2. Open an XLS (Excel 97-2003) file using the string constructor.
            // -----------------------------------------------------------------
            Workbook workbookXls = new Workbook(xlsPath);
            Console.WriteLine($"Opened XLS workbook. Worksheets count: {workbookXls.Worksheets.Count}");
            workbookXls.Save("output_from_xls.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Saved XLS workbook as XLSX.");

            // -----------------------------------------------------------------
            // 3. Open a CSV file using LoadOptions to specify the format explicitly.
            // -----------------------------------------------------------------
            Workbook workbookCsv = null;
            if (File.Exists(csvPath))
            {
                LoadOptions csvLoadOptions = new LoadOptions(LoadFormat.Csv);
                workbookCsv = new Workbook(csvPath, csvLoadOptions);
                Console.WriteLine($"Opened CSV workbook. Worksheets count: {workbookCsv.Worksheets.Count}");
                workbookCsv.Save("output_from_csv.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Saved CSV workbook as XLSX.");
            }
            else
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
            }

            // -----------------------------------------------------------------
            // 4. Open a JSON file using LoadOptions.
            // -----------------------------------------------------------------
            Workbook workbookJson = null;
            if (File.Exists(jsonPath))
            {
                LoadOptions jsonLoadOptions = new LoadOptions(LoadFormat.Json);
                workbookJson = new Workbook(jsonPath, jsonLoadOptions);
                Console.WriteLine($"Opened JSON workbook. Worksheets count: {workbookJson.Worksheets.Count}");
                workbookJson.Save("output_from_json.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Saved JSON workbook as XLSX.");
            }
            else
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
            }

            // -----------------------------------------------------------------
            // 5. Open an EPUB ebook using the dedicated EbookLoadOptions.
            // -----------------------------------------------------------------
            Workbook workbookEpub = null;
            if (File.Exists(epubPath))
            {
                EbookLoadOptions epubLoadOptions = new EbookLoadOptions(LoadFormat.Epub);
                workbookEpub = new Workbook(epubPath, epubLoadOptions);
                Console.WriteLine($"Opened EPUB workbook. Worksheets count: {workbookEpub.Worksheets.Count}");
                workbookEpub.Save("output_from_epub.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Saved EPUB workbook as XLSX.");
            }
            else
            {
                Console.WriteLine($"EPUB file not found: {epubPath}");
            }

            // -----------------------------------------------------------------
            // 6. Demonstrate opening a workbook from a memory stream.
            // -----------------------------------------------------------------
            Workbook tempWorkbook = new Workbook();
            tempWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from stream");
            using (MemoryStream ms = new MemoryStream())
            {
                tempWorkbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;

                Workbook workbookFromStream = new Workbook(ms);
                Console.WriteLine($"Opened workbook from stream. Worksheets count: {workbookFromStream.Worksheets.Count}");
                workbookFromStream.Save("output_from_stream.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Saved workbook loaded from stream as XLSX.");
                workbookFromStream.Dispose();
            }

            // -----------------------------------------------------------------
            // 7. Detect file format using FileFormatUtil before loading (optional).
            // -----------------------------------------------------------------
            if (File.Exists(csvPath))
            {
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(csvPath);
                Console.WriteLine($"Detected format for '{csvPath}': {info.FileFormatType}, LoadFormat: {info.LoadFormat}");
            }

            // Clean up.
            workbookXlsx.Dispose();
            workbookXls.Dispose();
            workbookCsv?.Dispose();
            workbookJson?.Dispose();
            workbookEpub?.Dispose();
            tempWorkbook.Dispose();

            Console.WriteLine("All operations completed successfully.");
        }
    }
}