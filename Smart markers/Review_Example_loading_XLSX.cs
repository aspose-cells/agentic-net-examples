using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxExample
    {
        public static void Run()
        {
            // Determine the directory of the executing assembly
            string dataDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string sourcePath = Path.Combine(dataDir, "example.xlsx");

            // If the source file does not exist, create a simple workbook for demonstration
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Name = "SampleSheet";
                tempSheet.Cells["A1"].PutValue(123);
                tempSheet.Cells["B1"].Formula = "=A1*2";
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
                Console.WriteLine($"Created sample workbook at: {sourcePath}");
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Display basic worksheet information
            Console.WriteLine("Worksheet Name: " + sheet.Name);
            Console.WriteLine("Rows with data: " + (sheet.Cells.MaxDataRow + 1));
            Console.WriteLine("Columns with data: " + (sheet.Cells.MaxDataColumn + 1));

            // Load the same file without parsing formulas
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };
            Workbook workbookNoFormula = new Workbook(sourcePath, loadOptions);
            Console.WriteLine("Loaded without parsing formulas. First cell formula: " +
                              workbookNoFormula.Worksheets[0].Cells["A1"].Formula);

            // Save the workbook to a new file
            string outputPath = Path.Combine(dataDir, "example_loaded.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadXlsxExample.Run();
        }
    }
}