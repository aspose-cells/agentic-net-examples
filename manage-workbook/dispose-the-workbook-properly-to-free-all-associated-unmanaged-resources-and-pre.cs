using System;
using Aspose.Cells;

namespace AsposeCellsDisposeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook instance (create rule)
            Workbook workbook = new Workbook();

            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data into the worksheet
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Save the workbook to a file (save rule)
            string outputPath = "DisposedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Dispose the workbook to release unmanaged resources
            workbook.Dispose();

            // Optionally, inform the user
            Console.WriteLine($"Workbook saved to '{outputPath}' and disposed successfully.");
        }
    }
}