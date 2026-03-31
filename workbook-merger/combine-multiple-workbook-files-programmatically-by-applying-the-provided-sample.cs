using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineMultipleWorkbooksDemo
    {
        public static void Run()
        {
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            Workbook destWorkbook = new Workbook();

            foreach (string filePath in sourceFiles)
            {
                Workbook sourceWorkbook = new Workbook(filePath);
                destWorkbook.Combine(sourceWorkbook);
            }

            string outputPath = "CombinedOutput.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks combined successfully. Output saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CombineMultipleWorkbooksDemo.Run();
        }
    }
}