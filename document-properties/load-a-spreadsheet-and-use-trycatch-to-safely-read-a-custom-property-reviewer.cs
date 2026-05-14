using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ReadReviewerProperty
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Attempt to read the custom document property named "Reviewer"
            // This may throw if the property does not exist or other issues occur
            string reviewer = workbook.CustomDocumentProperties["Reviewer"]?.Value?.ToString();

            if (reviewer != null)
            {
                Console.WriteLine($"Reviewer: {reviewer}");
            }
            else
            {
                Console.WriteLine("The custom property \"Reviewer\" was not found.");
            }
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine($"Aspose.Cells error (Code {ex.Code}): {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}