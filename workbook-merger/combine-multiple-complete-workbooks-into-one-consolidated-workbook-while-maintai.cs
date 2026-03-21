using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConsolidateWorkbooksDemo
    {
        public static void Run()
        {
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            Workbook consolidatedWorkbook = new Workbook();

            foreach (string filePath in sourceFiles)
            {
                Workbook sourceWorkbook = new Workbook(filePath);
                consolidatedWorkbook.Combine(sourceWorkbook);
            }

            string outputPath = "ConsolidatedWorkbook.xlsx";
            consolidatedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"All workbooks have been merged into '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConsolidateWorkbooksDemo.Run();
        }
    }
}