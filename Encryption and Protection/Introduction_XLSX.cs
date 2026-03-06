using System;
using Aspose.Cells;

namespace AsposeCellsIntroduction
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook instance (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the default first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to cells
            worksheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            worksheet.Cells["B1"].PutValue(DateTime.Now);
            worksheet.Cells["A2"].PutValue("Sample Number:");
            worksheet.Cells["B2"].PutValue(12345);

            // Save the workbook to an XLSX file (lifecycle save rule)
            string outputPath = "Introduction.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook created and saved to '{outputPath}'.");
        }
    }
}