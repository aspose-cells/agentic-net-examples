using System;
using System.IO;
using Aspose.Cells;

class ExtractTabIdDemo
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(filePath);

            // Iterate through each worksheet and output its TabId
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int tabId = sheet.TabId; // TabId property may be 0 if not explicitly set
                Console.WriteLine($"Sheet \"{sheet.Name}\" has TabId: {tabId}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}