using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxSaveMhtmlDemo
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Path for the output MHTML file
            string outputPath = "output.mhtml";

            // Save the workbook in MHTML format
            workbook.Save(outputPath, SaveFormat.MHtml);

            Console.WriteLine($"Workbook successfully converted from '{sourcePath}' to '{outputPath}'.");
        }
    }
}