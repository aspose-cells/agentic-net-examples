using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data in column A (first row is header)
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(20);
        cells["A4"].PutValue(30);
        // ... add more rows as needed

        // Define a dynamic named range that expands with non‑empty rows in column A
        // Formula: =OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)
        // Starts at A2, height = count of non‑blank cells in column A minus the header
        int nameIdx = workbook.Worksheets.Names.Add("DynamicRange");
        Name dynamicName = workbook.Worksheets.Names[nameIdx];
        dynamicName.RefersTo = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)";

        // Save the workbook (lifecycle save)
        workbook.Save("DynamicNamedRange.xlsx");
    }
}