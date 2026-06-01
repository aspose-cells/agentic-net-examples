using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Write header.
            sheet.Cells[0, 0].PutValue("Item");
            sheet.Cells[0, 1].PutValue("Quantity");

            // Write data rows.
            sheet.Cells[1, 0].PutValue("Apples");
            sheet.Cells[1, 1].PutValue(10);

            sheet.Cells[2, 0].PutValue("Oranges");
            sheet.Cells[2, 1].PutValue(15);

            // Set the default printer paper size for the workbook (A5).
            workbook.Settings.PaperSize = PaperSizeType.PaperA5;

            // Define output path and ensure the directory exists.
            string outputPath = "PrintedA5Workbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to an XLSX file.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}