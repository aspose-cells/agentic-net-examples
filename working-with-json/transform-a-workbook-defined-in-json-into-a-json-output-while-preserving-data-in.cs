using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonTransform
{
    class Program
    {
        static void Main()
        {
            string inputJsonPath = "inputWorkbook.json";
            string outputJsonPath = "outputWorkbook.json";

            if (!File.Exists(inputJsonPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells["B1"].PutValue(123);
                tempWb.Save(inputJsonPath, SaveFormat.Json);
            }

            // Load the workbook from JSON
            Workbook workbook = new Workbook(inputJsonPath);

            // Save the workbook back to JSON (you can customize save options if needed)
            workbook.Save(outputJsonPath, SaveFormat.Json);

            Console.WriteLine($"Workbook JSON transformed and saved to: {outputJsonPath}");
        }
    }
}