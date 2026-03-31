using System;
using Aspose.Cells;

namespace CombineWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Destination workbook (initially empty)
            Workbook combinedWorkbook = new Workbook();

            // Load each source workbook and combine it into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Merge the source workbook into the combined workbook
                combinedWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook as an XLSX file
            combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}