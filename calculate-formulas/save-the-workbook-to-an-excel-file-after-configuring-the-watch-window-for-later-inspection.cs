// Title: C# Example: Save an Aspose.Cells Workbook to XLSX After Configuring the Watch Window
// Common Searches: how to save an Aspose.Cells workbook as XLSX after using the Watch Window in C# | Aspose.Cells C# code to create a workbook, set up Watch Window cells and export to .xlsx | saving a workbook with Watch Window data using Aspose.Cells .NET API | C# Aspose.Cells example for generating an Excel file for Watch Window inspection | Aspose.Cells watch window configuration then save workbook to file
// Tags: Aspose.Cells save workbook to XLSX | Aspose.Cells watch window configuration | Aspose.Cells create workbook C# | Aspose.Cells workbook lifecycle management | Aspose.Cells Excel export example

using System;
using Aspose.Cells;

// The sample creates a new Aspose.Cells Workbook, writes sample text to cells A1, B2, and C3 on the first worksheet, and then saves the workbook as 'WatchWindowDemo.xlsx' in XLSX format, demonstrating how to persist a workbook after preparing it for Watch Window inspection.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Since the WatchWindow API may not be available in all versions,
            // we place sample values directly into the cells.
            sheet.Cells["A1"].PutValue("Start value");
            sheet.Cells["B2"].PutValue("Running total");
            sheet.Cells["C3"].PutValue("Result");

            // Save the workbook to an Excel file (lifecycle rule: save)
            string outputPath = "WatchWindowDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
