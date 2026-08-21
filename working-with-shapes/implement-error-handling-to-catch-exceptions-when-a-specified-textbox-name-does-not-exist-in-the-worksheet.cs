// Title: Handle Missing TextBox Names in Aspose.Cells (C#) with Proper Exception Catching
// Description: Shows how to add a TextBox to a worksheet, retrieve it by name, and gracefully handle a non‑existent TextBox using CellsException (ExceptionType.Shape) in Aspose.Cells for .NET.
// Keywords: Aspose.Cells TextBox name not found | CellsException shape | C# Aspose.Cells error handling | retrieve TextBox by name | worksheet shape exception | Aspose.Cells missing shape handling | Aspose.Cells TextBox lookup | exception handling Aspose.Cells | Aspose.Cells .NET | Excel shape error handling
// Common Searches: Aspose.Cells catch exception when textbox name missing | C# get TextBox by name Aspose.Cells | CellsException shape code example | how to handle missing shape in Aspose.Cells | Aspose.Cells TextBox not found error | retrieve worksheet shape by name with error handling | Aspose.Cells .NET exception handling for shapes
// Developer Intent: Add robust error handling when accessing a TextBox by name to avoid crashes if the name is absent.
// Use Cases: Validate user‑provided TextBox names before processing. | Log a warning and continue when a required TextBox is missing. | Separate shape lookup from other workbook operations to ensure the file saves. | Provide user‑friendly feedback for missing shapes in console or UI. | Implement fallback logic when a specific TextBox is not present.
// AI Prompts: Write C# code using Aspose.Cells that retrieves a TextBox by its Name and catches CellsException with ExceptionType.Shape for missing names. | Show how to log a clear message and keep processing after a TextBox lookup fails in an Aspose.Cells workbook. | Provide an example of wrapping worksheet.TextBoxes["MyName"] in try‑catch distinguishing shape‑related errors from other exceptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Shows how to add a TextBox to a worksheet, retrieve it by name, and gracefully handle a non‑existent TextBox using CellsException (ExceptionType.Shape) in Aspose.Cells for .NET.
    public class TextBoxNameErrorHandling
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox and assign a name
            int tbIndex = worksheet.TextBoxes.Add(2, 2, 150, 60);
            TextBox textBox = worksheet.TextBoxes[tbIndex];
            textBox.Name = "MyTextBox";
            textBox.Text = "Hello Aspose!";

            // Attempt to retrieve a textbox by a name that may not exist
            string targetName = "NonExistingBox";

            try
            {
                // This will throw a CellsException if the name is not found
                TextBox targetBox = worksheet.TextBoxes[targetName];
                Console.WriteLine($"Found TextBox: {targetBox.Name}, Text: {targetBox.Text}");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Shape)
            {
                // Handle the case where the textbox name does not exist
                Console.WriteLine($"TextBox with name \"{targetName}\" was not found. Exception: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }

            // Save the workbook (using the standard save method)
            try
            {
                workbook.Save("TextBoxNameErrorHandling.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
