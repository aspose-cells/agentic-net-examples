using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($" - {circularCellsData.Current}");
            }
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            string inputPath = "CircularSample.xlsx";

            if (!File.Exists(inputPath))
            {
                // Create a workbook with a circular reference for demonstration
                var wb = new Workbook();
                var sheet = wb.Worksheets[0];
                sheet.Cells["A1"].Formula = "=A2";
                sheet.Cells["A2"].Formula = "=A1";
                wb.Save(inputPath);
            }

            var workbook = new Workbook(inputPath);

            var calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            workbook.CalculateFormula(calcOptions);

            string outputPath = "CircularSample_Calculated.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Processing completed.");
        }
    }
}