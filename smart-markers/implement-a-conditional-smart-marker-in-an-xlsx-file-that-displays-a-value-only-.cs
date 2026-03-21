using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: A1 holds a numeric value, B1 holds the text to display conditionally
        sheet.Cells["A1"].PutValue(15);               // Change this value to test the condition
        sheet.Cells["B1"].PutValue("Show this text");

        // Smart marker in C1:
        // The expression displays the value of B1 only when A1 > 10; otherwise it shows an empty string.
        sheet.Cells["C1"].PutValue("&=IF(A1>10, B1, \"\")");

        // Process the smart markers in the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.Process();

        // Save the result to an XLSX file
        workbook.Save("ConditionalSmartMarker.xlsx");
    }
}