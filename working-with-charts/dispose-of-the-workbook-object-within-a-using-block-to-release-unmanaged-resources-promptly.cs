using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Workbook is created inside a using block so it will be disposed automatically
        using (Workbook workbook = new Workbook())
        {
            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some data to the worksheet
            sheet.Cells["A1"].PutValue("Hello, Aspose!");

            // Save the workbook to a file
            workbook.Save("DisposedWorkbook.xlsx", SaveFormat.Xlsx);
        } // workbook.Dispose() is called here
    }
}