using System;
using Aspose.Cells;

namespace WorkbookMergingConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the workbooks to be merged
            string sourcePath = "SourceWorkbook.xlsx";
            string destinationPath = "DestinationWorkbook.xlsx";
            string outputPath = "CombinedWorkbook.xlsx";

            // Create the source workbook (could also be loaded from an existing file)
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");
            sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);

            // Load the destination workbook from an existing file
            Workbook destWorkbook = new Workbook(destinationPath);
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }
    }
}