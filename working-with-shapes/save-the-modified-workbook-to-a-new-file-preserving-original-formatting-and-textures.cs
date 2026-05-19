using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file path)
            string sourcePath = "source.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify a cell value (example modification)
            sheet.Cells["A1"].PutValue("Modified Value");

            // Save the modified workbook to a new file while preserving all formatting and textures
            string outputPath = "modified_output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}