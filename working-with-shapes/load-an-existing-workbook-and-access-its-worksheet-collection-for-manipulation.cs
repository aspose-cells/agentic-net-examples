using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file to be loaded
            string inputPath = "input.xlsx";

            // Load the workbook from the file using the Workbook(string) constructor
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet collection of the loaded workbook
            WorksheetCollection worksheets = workbook.Worksheets;

            // Example manipulation: write a value to cell A1 of the first worksheet
            Worksheet firstSheet = worksheets[0];
            firstSheet.Cells["A1"].PutValue("Loaded and Modified");

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded from '{inputPath}', modified, and saved to '{outputPath}'.");
        }
    }
}