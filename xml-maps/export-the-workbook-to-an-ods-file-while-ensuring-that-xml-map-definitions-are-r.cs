using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorkbookToOds
    {
        public static void Run()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Save the workbook as ODS using default options
            string outputPath = "Workbook.ods";
            workbook.Save(outputPath, SaveFormat.ODS);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            ExportWorkbookToOds.Run();
        }
    }
}