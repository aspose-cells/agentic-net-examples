using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode for concurrent editing by multiple users
        workbook.Settings.Shared = true;

        // Optional: add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Shared Workbook");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Save the shared workbook
        string filePath = "SharedWorkbook.xlsx";
        workbook.Save(filePath);

        // Load the workbook to verify that the Shared property is set
        Workbook loadedWorkbook = new Workbook(filePath);
        Console.WriteLine("Shared property value: " + loadedWorkbook.Settings.Shared);
    }
}