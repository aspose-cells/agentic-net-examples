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

            // Retrieve the custom document property named "Reviewer"
            DocumentProperty reviewerProp = workbook.CustomDocumentProperties["Reviewer"];

            if (reviewerProp != null)
            {
                Console.WriteLine("Reviewer: " + reviewerProp.Value);
            }
            else
            {
                Console.WriteLine("Custom property 'Reviewer' not found.");
            }
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine("Aspose.Cells error: " + ex.Message);
            Console.WriteLine("Error code: " + ex.Code);
        }
        catch (Exception ex)
        {
            // Handle any other exceptions
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}