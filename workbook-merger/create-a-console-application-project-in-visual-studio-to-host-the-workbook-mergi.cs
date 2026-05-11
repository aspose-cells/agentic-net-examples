using System;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook(); // using Workbook() constructor
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (XLSX format) and add some data
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx); // using Workbook(FileFormatType) constructor
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook); // using Workbook.Combine method

            // Save the merged workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx); // using Workbook.Save method

            Console.WriteLine($"Workbooks merged successfully. Output file: {outputPath}");
        }
    }
}