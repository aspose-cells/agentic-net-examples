// Title: C# – Handle Missing TextBox Name Exception in Aspose.Cells Worksheets
// Description: Demonstrates how to safely retrieve a TextBox by name from a worksheet, catch the CellsException thrown when the specified name does not exist, provide a clear error message, and continue processing before saving the workbook.
// Keywords: Aspose.Cells | C# | TextBox | missing shape | exception handling | CellsException | shape lookup error | worksheet.TextBoxes indexer | Excel automation
// Common Searches: Aspose.Cells catch exception for non‑existent TextBox | C# retrieve TextBox by name error handling | Worksheet.TextBoxes name not found exception | How to handle missing shape in Aspose.Cells | CellsException Shape code example
// Developer Intent: The developer needs robust code that detects when a TextBox with a given name is absent in a worksheet and handles the situation without crashing the application.
// Use Cases: Validate user‑supplied TextBox identifiers before accessing them in a template workbook. | Log detailed information when a required shape is missing while processing Excel reports. | Display a friendly UI message instead of an unhandled exception when a TextBox cannot be found.
// AI Prompts: Create a C# helper method that returns null when worksheet.TextBoxes[name] throws a CellsException for a missing TextBox. | Show how to log CellsException details to a file and continue processing the workbook. | Write a custom exception filter that distinguishes missing shape names from other Aspose.Cells errors.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to safely retrieve a TextBox by name from a worksheet, catch the CellsException thrown when the specified name does not exist, provide a clear error message, and continue processing before saving the workbook.
class TextBoxNameExceptionDemo
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox and assign a name to it
        int textboxIndex = worksheet.TextBoxes.Add(1, 1, 150, 40);
        TextBox existingTextBox = worksheet.TextBoxes[textboxIndex];
        existingTextBox.Name = "MyBox";
        existingTextBox.Text = "Sample Text";

        // Name we want to retrieve – deliberately does not exist
        string targetName = "NonExistingBox";

        try
        {
            // Attempt to get the textbox by name
            TextBox targetTextBox = worksheet.TextBoxes[targetName];

            // If no exception, the textbox exists
            Console.WriteLine($"Found TextBox: {targetTextBox.Name}");
            Console.WriteLine($"Text: {targetTextBox.Text}");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Shape)
        {
            // Specific handling when the textbox name is not found
            Console.WriteLine($"TextBox with name '{targetName}' does not exist. Details: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General fallback for any other unexpected errors
            Console.WriteLine($"Error accessing TextBox '{targetName}': {ex.Message}");
        }

        // Save the workbook safely
        try
        {
            workbook.Save("TextBoxNameExceptionDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
