using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the custom document property named "Reviewer"
            DocumentProperty reviewerProp = workbook.CustomDocumentProperties["Reviewer"];

            if (reviewerProp != null)
            {
                // Property exists – output its value
                Console.WriteLine($"Reviewer: {reviewerProp.Value}");
            }
            else
            {
                // Property not found – inform the user
                Console.WriteLine("Custom property 'Reviewer' not found.");
            }
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors
            Console.WriteLine($"Aspose.Cells error: {ex.Message} (Code: {ex.Code})");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}