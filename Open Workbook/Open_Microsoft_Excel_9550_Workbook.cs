using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class OpenExcel95Workbook
    {
        static void Main()
        {
            // Path to the Excel 95 (BIFF5) workbook file
            string sourcePath = "Excel95Workbook.xls";

            // Ensure the file exists; if not, create a simple workbook and save it in Excel5 format
            if (!File.Exists(sourcePath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(sourcePath, SaveFormat.Excel97To2003);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Display basic information to confirm the file was opened
            Console.WriteLine("Workbook loaded successfully.");
            Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

            // Save the workbook in a modern format for further processing
            workbook.Save("ConvertedWorkbook.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as XLSX.");
        }
    }
}