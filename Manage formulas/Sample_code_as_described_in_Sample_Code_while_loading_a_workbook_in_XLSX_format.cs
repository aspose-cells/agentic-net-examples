using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxDemo
    {
        public static void Run()
        {
            // Directory of the executing assembly
            string dataDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputFile = Path.Combine(dataDir, "sample.xlsx");

            // If the input file does not exist, create a simple workbook for demonstration
            if (!File.Exists(inputFile))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Hello World");
                wb.Save(inputFile, SaveFormat.Xlsx);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Read a value from cell A1
            Console.WriteLine("Cell A1 value: " + worksheet.Cells["A1"].StringValue);

            // Modify a cell value
            worksheet.Cells["B2"].PutValue("Loaded with Aspose.Cells");

            // Save the workbook to a new file
            string outputFile = Path.Combine(dataDir, "sample_modified.xlsx");
            workbook.Save(outputFile, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded and saved successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadXlsxDemo.Run();
        }
    }
}