using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxWorkbookDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create LoadOptions – default options are sufficient for XLSX
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the constructor that accepts a file name and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: read the value of cell A1 and display it
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to a new file (demonstrates the save lifecycle rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded from '{sourcePath}' and saved to '{outputPath}'.");
        }
    }
}