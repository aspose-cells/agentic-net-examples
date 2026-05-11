using System;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the workbooks to be merged
            string firstWorkbookPath = "FirstWorkbook.xlsx";
            string secondWorkbookPath = "SecondWorkbook.xlsx";
            string outputPath = "CombinedWorkbook.xlsx";

            // Load the first workbook (this will be the destination workbook)
            Workbook destWorkbook = new Workbook(firstWorkbookPath);

            // Load the second workbook (source workbook)
            Workbook sourceWorkbook = new Workbook(secondWorkbookPath);

            // Combine the source workbook into the destination workbook.
            // Default Combine behavior preserves charts, images, and other objects.
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }
    }
}