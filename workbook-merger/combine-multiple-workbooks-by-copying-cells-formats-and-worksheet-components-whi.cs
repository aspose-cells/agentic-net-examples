using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineMultipleWorkbooks
    {
        public static void Run()
        {
            // Paths of the workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Ensure there is at least one workbook to start with
            if (sourceFiles.Length == 0)
            {
                Console.WriteLine("No source files provided.");
                return;
            }

            // Load the first workbook – it will act as the destination workbook
            Workbook destinationWorkbook = new Workbook(sourceFiles[0]);

            // Combine the remaining workbooks into the destination workbook
            for (int i = 1; i < sourceFiles.Length; i++)
            {
                // Load the next source workbook
                Workbook sourceWorkbook = new Workbook(sourceFiles[i]);

                // Merge the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook
            string outputPath = "CombinedWorkbook.xlsx";
            destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks combined successfully. Output saved to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CombineMultipleWorkbooks.Run();
        }
    }
}