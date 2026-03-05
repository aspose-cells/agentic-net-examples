using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsReviewExample
{
    class Program
    {
        static void Main()
        {
            string dataDir = AppDomain.CurrentDomain.BaseDirectory;
            string xlsxPath = Path.Combine(dataDir, "example.xlsx");
            string mhtmlPath = Path.Combine(dataDir, "output.mht");

            // Ensure the source XLSX file exists; if not, create a simple workbook.
            if (!File.Exists(xlsxPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Hello");
                ws.Cells["B2"].PutValue(123);
                wb.Save(xlsxPath, SaveFormat.Xlsx);
            }

            // Load the existing XLSX workbook.
            Workbook workbook = new Workbook(xlsxPath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Display basic information.
            Console.WriteLine("Worksheet Name: " + worksheet.Name);
            Console.WriteLine("Number of Cells: " + worksheet.Cells.Count);

            // Save the workbook in MHTML format.
            workbook.Save(mhtmlPath, SaveFormat.MHtml);

            // Load the MHTML file to verify conversion.
            Workbook workbookFromMhtml = new Workbook(mhtmlPath);

            // Output the number of worksheets in the MHTML file.
            Console.WriteLine("Number of Worksheets in MHTML: " + workbookFromMhtml.Worksheets.Count);
        }
    }
}