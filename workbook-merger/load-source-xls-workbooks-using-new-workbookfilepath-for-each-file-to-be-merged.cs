using System;
using Aspose.Cells;

namespace MergeWorkbooksExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths of the source XLS workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xls",
                "Source2.xls",
                "Source3.xls"
            };

            // Create a destination workbook (empty workbook)
            Workbook destinationWorkbook = new Workbook();

            // Load each source workbook using the Workbook(string) constructor
            foreach (string filePath in sourceFiles)
            {
                // Load source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the merged workbook to a new file
            destinationWorkbook.Save("MergedOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}