using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsLoadXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "sample.xlsx";

            // -------------------------------------------------
            // 1. Detect the file format using FileFormatUtil
            // -------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");

            // -------------------------------------------------
            // 2. Create LoadOptions with the detected format
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat);
            // Example: skip formula parsing to speed up loading
            loadOptions.ParsingFormulaOnOpen = false;
            // Example: do not check data validation while loading
            loadOptions.CheckDataValid = false;

            // -------------------------------------------------
            // 3. Load the workbook using the string + LoadOptions constructor
            // -------------------------------------------------
            Workbook workbook = new Workbook(sourcePath, loadOptions);
            Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");

            // -------------------------------------------------
            // 4. Access first worksheet and read a cell value
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine($"Cell A1 value: {cell.Value}");

            // -------------------------------------------------
            // 5. Save the workbook to a new file (demonstrates proper save usage)
            // -------------------------------------------------
            string outputPath = "output_copy.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");

            // -------------------------------------------------
            // 6. Optional: Convert the XLSX to another format using ConversionUtility
            // -------------------------------------------------
            string pdfPath = "output.pdf";
            ConversionUtility.Convert(outputPath, pdfPath);
            Console.WriteLine($"Workbook converted to PDF: {pdfPath}");
        }
    }
}