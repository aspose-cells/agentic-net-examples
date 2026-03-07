using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            string inputPath = "template.xlsx"; // path to the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Insert a smart marker using the {{field_name}} syntax into cell A1 of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("{{field_name}}");

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Smart marker inserted and workbook saved to '{outputPath}'.");
        }
    }
}