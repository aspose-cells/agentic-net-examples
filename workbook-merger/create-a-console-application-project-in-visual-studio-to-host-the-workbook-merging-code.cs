using System;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (empty) and add some data
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Save the merged workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Output file: {outputPath}");
        }
    }
}