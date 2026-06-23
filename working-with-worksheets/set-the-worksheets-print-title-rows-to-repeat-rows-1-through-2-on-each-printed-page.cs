using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetPrintTitleRowsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set rows 1 through 2 as print titles (repeat on each printed page)
            worksheet.PageSetup.PrintTitleRows = "$1:$2";

            // Define output file path
            string outputPath = "PrintTitleRowsRows1to2.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}