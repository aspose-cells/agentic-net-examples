using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be opened
            string inputPath = "input.xlsx";

            // Load the workbook (creates a new one if the file does not exist)
            Workbook workbook;
            if (System.IO.File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new empty workbook
            }

            // Add a custom document property named "ProcessedDate" with the current DateTime value
            DateTime now = DateTime.Now;
            workbook.CustomDocumentProperties.Add("ProcessedDate", now);

            // Save the workbook with the new property
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}