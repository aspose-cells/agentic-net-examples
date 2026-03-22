using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookCombineDemo
    {
        public static void Run()
        {
            // Create the first workbook and add some data
            Workbook workbook1 = new Workbook();
            Worksheet sheet1 = workbook1.Worksheets[0];
            sheet1.Name = "First";
            sheet1.Cells["A1"].PutValue("Data from Workbook 1");

            // Create the second workbook and add some data
            Workbook workbook2 = new Workbook();
            Worksheet sheet2 = workbook2.Worksheets[0];
            sheet2.Name = "Second";
            sheet2.Cells["A1"].PutValue("Data from Workbook 2");

            // Create the third workbook and add some data
            Workbook workbook3 = new Workbook();
            Worksheet sheet3 = workbook3.Worksheets[0];
            sheet3.Name = "Third";
            sheet3.Cells["A1"].PutValue("Data from Workbook 3");

            // Combine the second workbook into the first workbook
            workbook1.Combine(workbook2);

            // Combine the third workbook into the first workbook (which already contains workbook2)
            workbook1.Combine(workbook3);

            // Save the combined workbook to a file
            string outputPath = "CombinedWorkbook.xlsx";
            workbook1.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks combined and saved to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookCombineDemo.Run();
        }
    }
}