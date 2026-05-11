using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndSaveDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the first workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";
            sourceSheet.Cells["A1"].PutValue("Source Data");

            // Create the second workbook and add some data
            Workbook destinationWorkbook = new Workbook();
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            destSheet.Name = "Destination";
            destSheet.Cells["B2"].PutValue("Destination Data");

            // Combine the source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook);

            // Define the output file path
            string outputPath = "CombinedWorkbook.xlsx";

            // Save the merged workbook to the specified path in XLSX format
            destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved successfully to: {outputPath}");
        }
    }
}