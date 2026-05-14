using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Disable the LinksUpToDate property to prevent link checks
            workbook.BuiltInDocumentProperties.LinksUpToDate = false;

            // Save the workbook with the updated setting
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}