// Title: C# – Read the custom document property “Reviewer” from an Excel workbook with try‑catch using Aspose.Cells
// Description: Loads an Excel file into an Aspose.Cells Workbook, attempts to retrieve the custom document property named "Reviewer" via the CustomDocumentProperties collection, and safely handles missing properties, CellsException, and other exceptions with a try‑catch block.
// Keywords: Aspose.Cells | C# | .NET | Excel custom document property | Reviewer property | try‑catch | CellsException handling | load workbook | read custom property | error handling in Aspose.Cells
// Common Searches: how to read a custom document property reviewer using Aspose.Cells C# | Aspose.Cells try catch when custom property not found | C# read Excel custom property with error handling | retrieve custom document property Reviewer Aspose.Cells | handle CellsException reading document properties
// Developer Intent: Safely obtain the value of the "Reviewer" custom document property from an Excel file while handling missing properties and any Aspose.Cells or generic exceptions.
// Use Cases: Validate that a workbook includes a reviewer name before further processing. | Log or display reviewer information for audit trails. | Conditionally modify a workbook only when the Reviewer property exists.
// AI Prompts: Write C# code with Aspose.Cells that returns the value of a custom document property "Reviewer" or null, using try‑catch for CellsException and generic exceptions. | Show how to add a "Reviewer" custom property to a workbook if it is missing, then read and output its value in C#. | Create a reusable C# method that reads any specified custom document property from an Aspose.Cells workbook and gracefully handles missing properties and errors.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // Loads an Excel file into an Aspose.Cells Workbook, attempts to retrieve the custom document property named "Reviewer" via the CustomDocumentProperties collection, and safely handles missing properties, CellsException, and other exceptions with a try‑catch block.
    class ReadReviewerProperty
    {
        static void Main()
        {
            // Path to the Excel file
            string filePath = "input.xlsx";

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(filePath);

            try
            {
                // Attempt to read the custom document property named "Reviewer"
                DocumentProperty reviewerProp = workbook.CustomDocumentProperties["Reviewer"];

                // If the property does not exist, the indexer returns null
                if (reviewerProp == null)
                {
                    Console.WriteLine("Custom property 'Reviewer' not found.");
                }
                else
                {
                    Console.WriteLine($"Reviewer: {reviewerProp.Value}");
                }
            }
            catch (CellsException ex)
            {
                // Handle Aspose.Cells specific exceptions
                Console.WriteLine($"CellsException caught. Code: {ex.Code}, Message: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
