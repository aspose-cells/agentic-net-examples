using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteColumnAndRecalculateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in columns A and B
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["A2"].PutValue(30);
                cells["B2"].PutValue(40);

                // Add a formula that sums the values in column A and B
                cells["C1"].Formula = "=SUM(A1:B2)";

                // Calculate formulas before deletion (optional, just to show initial state)
                workbook.CalculateFormula();

                Console.WriteLine("Before deleting column:");
                Console.WriteLine($"C1 formula result: {cells["C1"].Value}");

                // Delete column B (index 1) and update references in other worksheets
                cells.DeleteColumn(1, true);

                // Recalculate formulas after the column deletion
                workbook.CalculateFormula();

                Console.WriteLine("After deleting column B:");
                Console.WriteLine($"C1 formula: {cells["C1"].Formula}");
                Console.WriteLine($"C1 formula result: {cells["C1"].Value}");

                // Define output file path
                string outputPath = "DeleteColumnAndRecalculateDemo.xlsx";

                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteColumnAndRecalculateDemo.Run();
        }
    }
}