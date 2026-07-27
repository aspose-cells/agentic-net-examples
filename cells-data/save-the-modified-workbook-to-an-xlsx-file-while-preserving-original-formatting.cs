using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file path)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Example modification: change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified Value");

            // Save the modified workbook to a new XLSX file while preserving all original formatting
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}