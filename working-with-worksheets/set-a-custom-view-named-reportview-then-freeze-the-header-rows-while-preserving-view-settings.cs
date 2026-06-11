using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Preserve the current view type (if you need to keep it unchanged)
            ViewType currentView = sheet.ViewType;

            // NOTE: The CustomViews collection is not available in the current Aspose.Cells version.
            // If needed, custom view functionality can be added in a later version.
            // The following code is kept for reference:
            // int viewIndex = sheet.CustomViews.Add("ReportView");

            // Freeze the header rows (freeze the first row, no frozen columns)
            // Row index is 1 (second row) because the first row (index 0) will be frozen
            sheet.FreezePanes(1, 0, 1, 0);

            // Restore the original view type to ensure view settings are preserved
            sheet.ViewType = currentView;

            // Define output file path
            string outputPath = "ReportViewDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}