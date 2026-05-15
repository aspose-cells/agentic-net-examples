using System;
using Aspose.Cells;

class SharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode for concurrent editing by multiple users
        workbook.Settings.Shared = true;

        // Add some sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Shared Workbook");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Save the workbook
        string outputPath = "SharedWorkbook.xlsx";
        workbook.Save(outputPath);

        // Load the workbook to verify the Shared setting
        Workbook loadedWorkbook = new Workbook(outputPath);
        Console.WriteLine("Shared property value: " + loadedWorkbook.Settings.Shared);
    }
}