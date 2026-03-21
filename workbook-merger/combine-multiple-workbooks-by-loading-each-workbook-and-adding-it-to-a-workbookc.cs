using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CombineMultipleWorkbooks.Run();
        }
    }

    public class CombineMultipleWorkbooks
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

            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Combined workbook saved to '{outputPath}'.");
        }
    }
}