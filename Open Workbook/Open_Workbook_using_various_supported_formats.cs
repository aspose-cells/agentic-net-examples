using System;
using System.IO;
using Aspose.Cells;

class OpenWorkbookDemo
{
    static void Main()
    {
        // Example 1: Open a workbook directly from a file path (XLSX)
        string xlsxPath = "sample.xlsx";
        Workbook wbFromFile = new Workbook(xlsxPath); // Workbook(string)
        Console.WriteLine($"Opened '{xlsxPath}' with format: {wbFromFile.FileFormat}");

        // Example 2: Open a workbook from a memory stream (CSV)
        string csvContent = "Name,Age\nJohn,30\nJane,25";
        byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(csvContent);
        using (MemoryStream csvStream = new MemoryStream(csvBytes))
        {
            Workbook wbFromStream = new Workbook(csvStream); // Workbook(Stream)
            Console.WriteLine($"Opened CSV stream with format: {wbFromStream.FileFormat}");
        }

        // Example 3: Open a workbook with explicit LoadOptions (XLSX)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook wbWithOptions = new Workbook(xlsxPath, loadOptions); // Workbook(string, LoadOptions)
        Console.WriteLine($"Opened with LoadOptions, format: {wbWithOptions.FileFormat}");

        // Example 4: Detect file format before loading
        FileFormatInfo info = FileFormatUtil.DetectFileFormat(xlsxPath); // DetectFileFormat(string)
        Console.WriteLine($"Detected format for '{xlsxPath}': {info.FileFormatType}");

        // Example 5: Open a workbook from a stream with LoadOptions (TSV)
        string tsvContent = "Product\tPrice\nApple\t1.2\nBanana\t0.8";
        byte[] tsvBytes = System.Text.Encoding.UTF8.GetBytes(tsvContent);
        using (MemoryStream tsvStream = new MemoryStream(tsvBytes))
        {
            LoadOptions tsvOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook wbTsv = new Workbook(tsvStream, tsvOptions); // Workbook(Stream, LoadOptions)
            Console.WriteLine($"Opened TSV stream with format: {wbTsv.FileFormat}");
        }

        // Example 6: Save the first workbook to PDF using Save(string, SaveFormat)
        wbFromFile.Save("output.pdf", SaveFormat.Pdf);
        Console.WriteLine("Saved workbook as PDF.");
    }
}