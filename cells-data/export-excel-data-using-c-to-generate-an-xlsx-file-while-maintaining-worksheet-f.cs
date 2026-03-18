using System;
using Aspose.Cells;

namespace AsposeCellsExportExample
{
    public class ExportExcel
    {
        public static void Run()
        {
            // Path to the source Excel file (any existing format, e.g., .xls, .csv, etc.)
            string sourcePath = "source.xlsx";

            // Load the workbook from the source file (preserves all worksheets, styles, formulas, etc.)
            Workbook workbook = new Workbook(sourcePath);

            // Define the output file path for the exported XLSX file
            string outputPath = "exported_output.xlsx";

            // Save the workbook as XLSX while keeping full worksheet fidelity
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportExcel.Run();
        }
    }
}