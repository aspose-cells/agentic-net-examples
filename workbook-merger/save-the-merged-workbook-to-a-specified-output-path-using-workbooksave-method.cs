using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the source workbooks and the final merged file
            string sourcePath1 = "Source1.xlsx";
            string sourcePath2 = "Source2.xlsx";
            string mergedOutputPath = "MergedWorkbook.xlsx";

            // -------------------------------------------------
            // Create first source workbook and add sample data
            // -------------------------------------------------
            Workbook sourceWorkbook1 = new Workbook();
            Worksheet sheet1 = sourceWorkbook1.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Data from first workbook");
            // Save to disk so it can be loaded later (optional)
            sourceWorkbook1.Save(sourcePath1, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Create second source workbook and add sample data
            // -------------------------------------------------
            Workbook sourceWorkbook2 = new Workbook();
            Worksheet sheet2 = sourceWorkbook2.Worksheets[0];
            sheet2.Name = "SecondSheet";
            sheet2.Cells["A1"].PutValue("Data from second workbook");
            sourceWorkbook2.Save(sourcePath2, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Load the two workbooks (demonstrates loading rule)
            // -------------------------------------------------
            Workbook wb1 = new Workbook(sourcePath1);
            Workbook wb2 = new Workbook(sourcePath2);

            // -------------------------------------------------
            // Combine the second workbook into the first one
            // -------------------------------------------------
            wb1.Combine(wb2);

            // -------------------------------------------------
            // Save the merged workbook to the specified output path
            // -------------------------------------------------
            wb1.Save(mergedOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved to: {Path.GetFullPath(mergedOutputPath)}");
        }
    }
}