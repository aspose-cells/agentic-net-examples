using System;
using Aspose.Cells;

public class WorkbookFileDemo
{
    public static void Main()
    {
        WriteWorkbookToFile();
    }

    public static void WriteWorkbookToFile()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Define the file name that will be saved
        string fileName = "Sample.xlsx";

        // Save the workbook to the specified file (format inferred from extension)
        workbook.Save(fileName);
    }
}