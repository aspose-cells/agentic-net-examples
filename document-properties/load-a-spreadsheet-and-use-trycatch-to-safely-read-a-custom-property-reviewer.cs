using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook (using the load rule)
        Workbook workbook = new Workbook(filePath);

        try
        {
            // Try to retrieve the custom document property named "Reviewer"
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
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine($"Aspose.Cells error: {ex.Message} (Code: {ex.Code})");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}