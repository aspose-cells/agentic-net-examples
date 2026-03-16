using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToOdsConversion
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the ODS output file
            string destPath = "output.ods";

            // Load the XLSX workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook in ODS format
            workbook.Save(destPath, SaveFormat.ODS);

            Console.WriteLine($"Conversion completed successfully: {sourcePath} -> {destPath}");
        }
    }
}