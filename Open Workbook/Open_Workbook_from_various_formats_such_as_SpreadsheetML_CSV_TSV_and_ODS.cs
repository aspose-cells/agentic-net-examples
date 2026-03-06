using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOpenVariousFormats
{
    class Program
    {
        static void Main()
        {
            // Paths to sample files in different formats.
            // Ensure these files exist in the execution directory or provide full paths.
            string spreadsheetMlPath = "sample.xml";   // SpreadsheetML (Excel 2003 XML)
            string csvPath = "sample.csv";            // CSV
            string tsvPath = "sample.tsv";            // TSV
            string odsPath = "sample.ods";            // OpenDocument Spreadsheet

            // Open SpreadsheetML using the simple string constructor.
            // Aspose.Cells automatically detects the format based on the file extension.
            Workbook workbookXml = new Workbook(spreadsheetMlPath);
            Console.WriteLine($"SpreadsheetML loaded. Worksheets count: {workbookXml.Worksheets.Count}");
            // Save as XLSX to verify successful load.
            workbookXml.Save("ConvertedFromSpreadsheetML.xlsx", SaveFormat.Xlsx);

            // Open CSV with explicit LoadOptions specifying LoadFormat.Csv.
            LoadOptions csvLoadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbookCsv = new Workbook(csvPath, csvLoadOptions);
            Console.WriteLine($"CSV loaded. Worksheets count: {workbookCsv.Worksheets.Count}");
            workbookCsv.Save("ConvertedFromCsv.xlsx", SaveFormat.Xlsx);

            // Open TSV with explicit LoadOptions specifying LoadFormat.Tsv.
            LoadOptions tsvLoadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbookTsv = new Workbook(tsvPath, tsvLoadOptions);
            Console.WriteLine($"TSV loaded. Worksheets count: {workbookTsv.Worksheets.Count}");
            workbookTsv.Save("ConvertedFromTsv.xlsx", SaveFormat.Xlsx);

            // Open ODS (OpenDocument Spreadsheet) using the simple string constructor.
            Workbook workbookOds = new Workbook(odsPath);
            Console.WriteLine($"ODS loaded. Worksheets count: {workbookOds.Worksheets.Count}");
            workbookOds.Save("ConvertedFromOds.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbookXml.Dispose();
            workbookCsv.Dispose();
            workbookTsv.Dispose();
            workbookOds.Dispose();

            Console.WriteLine("All files processed successfully.");
        }
    }
}