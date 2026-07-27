using System;
using Aspose.Cells;

class StrictComplianceSaveDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with a valid file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Enable ISO/IEC 29500:2008 Strict compliance for OOXML
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Define a path that typically requires elevated permissions
        string restrictedPath = @"C:\Windows\System32\restricted.xlsx";

        try
        {
            // Attempt to save the workbook to the restricted location
            workbook.Save(restrictedPath);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (UnauthorizedAccessException ex)
        {
            // Expected when the application lacks write permission
            Console.WriteLine("Failed to save workbook due to insufficient permissions: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine("An error occurred while saving the workbook: " + ex.Message);
        }
    }
}