using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDeveloperGuide
{
    class Program
    {
        static void Main()
        {
            // Directory for sample files.
            string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
            Directory.CreateDirectory(dataDir);

            // Source Excel file path.
            string sourceFile = Path.Combine(dataDir, "Sample.xlsx");

            // If the source file does not exist, create a simple workbook.
            if (!File.Exists(sourceFile))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Hello Aspose!");
                wb.Save(sourceFile, SaveFormat.Xlsx);
                wb.Dispose();
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(sourceFile);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"Loaded worksheet name: {sheet.Name}");
            Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

            // Save as PDF.
            string pdfFile = Path.Combine(dataDir, "SampleConverted.pdf");
            workbook.Save(pdfFile, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to: {pdfFile}");

            // Save a copy in XLSX format (format inferred from extension).
            string copyFile = Path.Combine(dataDir, "SampleCopy.xlsx");
            workbook.Save(copyFile);
            Console.WriteLine($"Workbook saved as a copy to: {copyFile}");

            // Clean up.
            workbook.Dispose();
        }
    }
}