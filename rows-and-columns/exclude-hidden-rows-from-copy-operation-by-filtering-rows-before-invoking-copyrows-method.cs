using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CopyVisibleRowsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cells destinationCells = destinationSheet.Cells;

            int destRowIndex = 0;
            int totalRows = sourceCells.MaxDataRow + 1; // MaxDataRow is zero‑based

            // Copy only visible rows
            for (int srcRowIndex = 0; srcRowIndex < totalRows; srcRowIndex++)
            {
                if (sourceCells.IsRowHidden(srcRowIndex))
                    continue;

                destinationCells.CopyRows(sourceCells, srcRowIndex, destRowIndex, 1);
                destRowIndex++;
            }

            // Save the result
            destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}