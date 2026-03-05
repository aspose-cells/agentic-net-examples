using System;
using Aspose.Cells;

namespace AsposeCellsSample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file to be loaded
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file using the string constructor
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Write some data to cells
            worksheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            worksheet.Cells["B1"].PutValue(DateTime.Now);

            // Optionally, you can modify other properties, e.g., set the default font
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // Save the modified workbook to a new file in XLSX format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Inform the user that the operation completed
            Console.WriteLine($"Workbook loaded from '{inputPath}', modified, and saved as '{outputPath}'.");
        }
    }
}