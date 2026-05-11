using System;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create source workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create destination workbook (empty or existing)
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks combined and saved to {outputPath}");
        }
    }
}