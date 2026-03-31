using System;
using Aspose.Cells;

public class WorksheetVisibilityDemo
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook (contains a default worksheet)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet is displayed (visible) in the workbook
        worksheet.IsVisible = true; // or worksheet.Visibility = SheetVisibility.Visible;

        // Save the workbook to a file
        string outputPath = "VisibleSheetWorkbook.xlsx";
        workbook.Save(outputPath);
    }
}