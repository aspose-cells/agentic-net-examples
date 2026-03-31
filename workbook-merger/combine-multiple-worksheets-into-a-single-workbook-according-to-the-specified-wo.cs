using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineWorksheetsDemo
    {
        public static void Run()
        {
            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Define the source workbook file paths to be merged
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

            // Load each source workbook and combine its worksheets into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Merge all worksheets from the source workbook into the destination workbook
                destWorkbook.Combine(sourceWorkbook);
            }

            // Save the resulting workbook that now contains all merged worksheets
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CombineWorksheetsDemo.Run();
        }
    }
}