// Title: Read the custom 'Reviewer' property from an Excel workbook with Aspose.Cells in C# using try‑catch for safe access
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells and retrieves the custom property named 'Reviewer' inside a try‑catch block. | Show how to verify the presence of a custom document property before reading it and output a fallback message when it does not exist using Aspose.Cells. | Demonstrate robust error handling while accessing Excel custom metadata with Aspose.Cells, including logging the caught exception message.
// Common Searches: aspocells c# how to safely get custom document property reviewer from excel workbook | c# exception handling reading custom metadata with Aspose.Cells | example code for checking existence of a custom attribute in .xlsx using Aspose.Cells | handle missing reviewer metadata in Aspose.Cells workbook without crashing
// Tags: Aspose.Cells custom metadata extraction | C# exception handling for Excel workbook properties | verify existence of custom field Aspose.Cells | read reviewer value from .xlsx file | safe workbook property access with Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an Excel workbook via Aspose.Cells, then uses a try‑catch block to safely check for and read the custom document property named 'Reviewer', printing its value or handling any errors gracefully.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        string reviewer = string.Empty;

        try
        {
            // Access the custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Attempt to read the "Reviewer" property safely
            if (customProps.Contains("Reviewer"))
            {
                // Retrieve the value and convert to string
                reviewer = customProps["Reviewer"].Value?.ToString();
                Console.WriteLine($"Reviewer: {reviewer}");
            }
            else
            {
                Console.WriteLine("Custom property 'Reviewer' not found.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during property access
            Console.WriteLine($"Error reading custom property: {ex.Message}");
        }
    }
}
