using System;
using Aspose.Cells;

class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "example.xlsx";

        try
        {
            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook opened successfully by checking the worksheet count
            Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            // If an exception occurs, the workbook failed to load
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
        }
    }
}