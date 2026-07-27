using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to column A (index 0)
        worksheet.Cells["A1"].PutValue("Short");
        worksheet.Cells["A2"].PutValue("A bit longer text");
        worksheet.Cells["A3"].PutValue("The longest entry in this column for testing auto‑fit");
        worksheet.Cells["A4"].PutValue("Mid");
        worksheet.Cells["A5"].PutValue("Another entry");

        // Auto‑fit only column A (column index 0)
        worksheet.AutoFitColumn(0);

        // Save the workbook
        workbook.Save("AutoFitColumnDemo.xlsx");
    }
}

// Author: Aspose.Cells example code.