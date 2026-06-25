using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to source workbooks
            string sourcePath1 = "SourceWorkbook1.xlsx";
            string sourcePath2 = "SourceWorkbook2.xlsx";

            // Load source workbooks
            Workbook sourceWb1 = new Workbook(sourcePath1);
            Workbook sourceWb2 = new Workbook(sourcePath2);

            // Create a new target workbook (initially contains one default sheet)
            Workbook targetWb = new Workbook();

            // -------------------------------------------------
            // Copy specific worksheets from the first source
            // -------------------------------------------------
            // Example: copy worksheet named "Data" from sourceWb1
            Worksheet sourceSheet1 = sourceWb1.Worksheets["Data"];
            if (sourceSheet1 != null)
            {
                // Add a new blank worksheet to the target workbook
                Worksheet destSheet1 = targetWb.Worksheets.Add("DataCopy");

                // Copy contents and formats from the source worksheet
                destSheet1.Copy(sourceSheet1);
            }

            // -------------------------------------------------
            // Copy specific worksheets from the second source
            // -------------------------------------------------
            // Example: copy worksheet at index 0 (first sheet) from sourceWb2
            Worksheet sourceSheet2 = sourceWb2.Worksheets[0];
            if (sourceSheet2 != null)
            {
                // Add a new blank worksheet to the target workbook
                Worksheet destSheet2 = targetWb.Worksheets.Add("FirstSheetCopy");

                // Copy contents and formats from the source worksheet
                destSheet2.Copy(sourceSheet2);
            }

            // -------------------------------------------------
            // Save the target workbook
            // -------------------------------------------------
            string outputPath = "CombinedWorkbook.xlsx";
            targetWb.Save(outputPath);

            Console.WriteLine($"Worksheets copied successfully. Output file: {outputPath}");
        }
    }
}