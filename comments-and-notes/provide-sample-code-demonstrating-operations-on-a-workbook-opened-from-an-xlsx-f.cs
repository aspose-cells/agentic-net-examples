using System;
using Aspose.Cells;

namespace AsposeCellsXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read existing value from cell A1
            Console.WriteLine("Original A1 value: " + sheet.Cells["A1"].StringValue);

            // Write new data to some cells
            sheet.Cells["A2"].PutValue("Aspose");
            sheet.Cells["B2"].PutValue(DateTime.Now);
            sheet.Cells["C2"].PutValue(12345);

            // Add a new worksheet and populate it
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            summarySheet.Cells["A1"].PutValue("Report generated on:");
            summarySheet.Cells["B1"].PutValue(DateTime.Now);

            // Save the modified workbook back to disk
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}