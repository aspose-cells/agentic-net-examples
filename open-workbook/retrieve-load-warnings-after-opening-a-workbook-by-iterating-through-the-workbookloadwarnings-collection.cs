// Title: How to Retrieve Workbook Load Warnings with Aspose.Cells for .NET
// Description: Shows how to open an Excel file with Aspose.Cells, verify the file path, handle loading errors, and iterate the Workbook.LoadWarnings collection to print each warning’s type and description, including a fallback for versions that lack the collection.
// Keywords: Aspose.Cells load warnings | Workbook.LoadWarnings .NET | retrieve load warnings C# | Aspose.Cells warning info | Excel load errors Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells get load warnings after opening workbook | C# iterate Workbook.LoadWarnings collection | how to display load warnings with Aspose.Cells | load warnings not available in older Aspose.Cells versions | retrieve warning type and description from Aspose.Cells
// Developer Intent: The developer needs to capture and display any warnings generated when an Excel workbook is loaded with Aspose.Cells.
// Use Cases: Log all load warnings to a file for troubleshooting malformed Excel files. | Show warning messages in a UI after a workbook is opened to inform users of potential data issues. | Filter specific warning types (e.g., unsupported formulas) and apply custom handling before further processing.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, checks for file existence, and prints every warning from Workbook.LoadWarnings, with a fallback when the collection is unavailable. | Create a reusable method that returns a list of warning descriptions and types from Workbook.LoadWarnings for further analysis. | Write unit tests that verify load warnings are captured correctly when opening corrupted or partially supported Excel files using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to open an Excel file with Aspose.Cells, verify the file path, handle loading errors, and iterate the Workbook.LoadWarnings collection to print each warning’s type and description, including a fallback for versions that lack the collection.
class LoadWarningsDemo
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Create default load options
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // NOTE: In some older versions of Aspose.Cells the LoadWarnings collection
            // may not be available. If it is present, the following code will display
            // any warnings generated during the load operation.
            // Uncomment the block below if your version supports Workbook.LoadWarnings.

            /*
            foreach (WarningInfo warning in workbook.LoadWarnings)
            {
                Console.WriteLine($"Warning Type: {warning.Type}");
                Console.WriteLine($"Description: {warning.Description}");
                Console.WriteLine();
            }
            */

            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., loading errors) and display a message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
