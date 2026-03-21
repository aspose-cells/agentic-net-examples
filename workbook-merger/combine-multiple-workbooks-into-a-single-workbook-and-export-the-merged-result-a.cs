using System;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Iterate through each source file, load it and combine with the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the merged workbook as an XLSX file
            destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}