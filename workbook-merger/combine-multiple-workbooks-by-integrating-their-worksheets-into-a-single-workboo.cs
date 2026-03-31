using System;
using Aspose.Cells;

namespace AsposeCellsCombineExample
{
    class Program
    {
        static void Main()
        {
            // Paths of source workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Destination workbook – start with an empty workbook (default Xlsx format)
            Workbook destWorkbook = new Workbook();

            // Iterate through each source file, load it, and combine its worksheets into the destination
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook srcWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                // This merges all worksheets, preserving data, styles, formulas, etc.
                destWorkbook.Combine(srcWorkbook);
            }

            // Save the combined workbook to a new file
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"All workbooks merged successfully. Output saved to: {outputPath}");
        }
    }
}