using System;
using Aspose.Cells;

namespace AsposeCellsLoadXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Detect the file format of the source file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");
            Console.WriteLine($"Detected FileFormatType: {formatInfo.FileFormatType}");

            // Create LoadOptions using the detected LoadFormat
            LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat);

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);
            Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");

            // Example operation: write a timestamp into the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue($"Loaded on {DateTime.Now}");

            // Save the workbook to a new file (still XLSX format)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as: {outputPath}");
        }
    }
}