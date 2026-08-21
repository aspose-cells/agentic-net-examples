// Title: C# – Read the custom document property “Reviewer” from an Excel workbook using Aspose.Cells with robust try‑catch handling
// Description: Load an Excel file with Aspose.Cells, retrieve the custom document property named "Reviewer", display its value or a not‑found message, and safely manage Aspose.Cells‑specific and generic exceptions using a try‑catch block.
// Keywords: Aspose.Cells | C# | custom document property | Reviewer | read Excel property | try catch | CellsException | error handling | Workbook loading | document properties
// Common Searches: Aspose.Cells read custom property Reviewer C# | how to get Excel custom document property with Aspose.Cells | C# try catch example for Aspose.Cells workbook loading | catch CellsException when reading Excel properties | check if custom property exists in Excel using Aspose
// Developer Intent: Load a workbook and safely obtain the "Reviewer" custom property while handling possible Aspose.Cells and general runtime errors.
// Use Cases: Validate that a spreadsheet contains a reviewer name before generating a report. | Log reviewer information for audit trails when processing batches of workbooks. | Provide a clear user message when the "Reviewer" property is missing.
// AI Prompts: Write C# code that uses Aspose.Cells to read a custom document property called "Reviewer" with proper try‑catch handling for CellsException and generic exceptions. | Explain how to add a custom document property "Reviewer" to a workbook and then read it safely in C# using Aspose.Cells. | Refactor the sample to log errors to a file instead of writing them to the console.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Load an Excel file with Aspose.Cells, retrieve the custom document property named "Reviewer", display its value or a not‑found message, and safely manage Aspose.Cells‑specific and generic exceptions using a try‑catch block.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Attempt to retrieve the custom document property named "Reviewer"
            DocumentProperty reviewerProperty = workbook.CustomDocumentProperties["Reviewer"];

            if (reviewerProperty != null)
            {
                // Property exists – output its value
                Console.WriteLine($"Reviewer: {reviewerProperty.Value}");
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
