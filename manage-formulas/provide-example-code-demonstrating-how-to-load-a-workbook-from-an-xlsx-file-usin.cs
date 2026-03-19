using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    public class LoadWorkbookDemo
    {
        public static void Main()
        {
            // Path to the existing XLSX file
            string inputPath = "sample.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read and display the value of cell A1
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Optionally, save the workbook to a new file to demonstrate the save rule
            string outputPath = "loaded_copy.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded and saved successfully.");
        }
    }
}