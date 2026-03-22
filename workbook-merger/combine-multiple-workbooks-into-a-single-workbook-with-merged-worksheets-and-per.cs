using System;
using Aspose.Cells;

namespace CombineWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Iterate through each source workbook, load it and combine into the destination
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Merge the source workbook's worksheets into the destination workbook
                destWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook as XLSX
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}