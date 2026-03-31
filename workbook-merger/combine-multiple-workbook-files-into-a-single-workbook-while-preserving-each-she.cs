using System;
using Aspose.Cells;

namespace AsposeCellsCombineExample
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Load each source workbook and combine it into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Merge the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook to disk (preserves all sheets, data, and formatting)
            destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}