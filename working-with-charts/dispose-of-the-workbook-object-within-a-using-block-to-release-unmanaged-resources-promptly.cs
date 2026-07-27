using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the Workbook inside a using block so it is disposed automatically
        using (Workbook workbook = new Workbook())
        {
            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Save the workbook to a file (uses the Save(string, SaveFormat) rule)
            workbook.Save("DisposedWorkbook.xlsx", SaveFormat.Xlsx);
        } // workbook.Dispose() is called here, releasing unmanaged resources
    }
}