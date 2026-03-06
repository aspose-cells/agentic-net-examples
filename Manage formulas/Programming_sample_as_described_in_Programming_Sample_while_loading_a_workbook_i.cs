using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSample
{
    public class LoadAndSaveXlsx
    {
        public static void Run()
        {
            // Path to the directory containing the Excel file.
            string dataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "YourDocumentDirectory");
            Directory.CreateDirectory(dataDir);

            string inputPath = Path.Combine(dataDir, "input.xlsx");
            string outputPath = Path.Combine(dataDir, "output.xlsx");

            // If the input file does not exist, create a simple workbook to work with.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(inputPath);
            }

            // Load the existing XLSX workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Read the current value of cell A1.
            Console.WriteLine("Original A1 value: " + worksheet.Cells["A1"].StringValue);

            // Update the value of cell A1.
            worksheet.Cells["A1"].PutValue("Updated Value");

            // Optionally, add a new row of data.
            worksheet.Cells["A2"].PutValue("Sample");
            worksheet.Cells["B2"].PutValue(12345);

            // Save the modified workbook.
            workbook.Save(outputPath);

            Console.WriteLine("Workbook loaded, modified, and saved successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadAndSaveXlsx.Run();
        }
    }
}