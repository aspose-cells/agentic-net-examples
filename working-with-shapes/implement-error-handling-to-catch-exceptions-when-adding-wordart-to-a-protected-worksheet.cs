// Title: C# example: Catch exceptions when adding WordArt (TextEffect) to a protected worksheet using Aspose.Cells
// AI Prompts: Write C# code that protects an Excel worksheet, attempts to insert a WordArt TextEffect shape with Aspose.Cells, and captures any protection‑related exceptions. | Show how to structure nested try‑catch blocks to handle errors both for shape insertion and workbook saving when using Aspose.Cells. | Provide a snippet that logs the exception message when adding a TextEffect to a locked sheet and continues execution.
// Common Searches: Aspose.Cells C# add WordArt to a locked worksheet and handle exception | how to catch protection error when inserting TextEffect shape with Aspose.Cells | C# try‑catch example for adding WordArt to a protected Excel sheet using Aspose.Cells
// Tags: add wordart to protected sheet aspocells | aspocells shape insertion exception handling | protect excel sheet before adding shape c# | c# aspocells workbook save error handling | text effect shape on protected sheet aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for shape operations

// The sample creates a workbook, protects the first worksheet, then tries to add a WordArt (TextEffect) shape inside a try‑catch block that captures protection‑related errors. It also demonstrates separate error handling for saving the workbook, with an outer catch for any unexpected exceptions.
class WordArtProtectedSheetExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet (no password for simplicity)
            sheet.Protect(ProtectionType.All);

            // Attempt to add WordArt (TextEffect) to the protected worksheet
            try
            {
                // Add WordArt with specified parameters
                // Parameters: preset, text, font name, font size, isBold, isItalic,
                // left, top, width, height, shapeRotation, textRotation
                sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Aspose.Cells",
                    "Arial",
                    36,
                    false,
                    false,
                    100,
                    100,
                    300,
                    100,
                    0,
                    0);
                Console.WriteLine("WordArt added successfully.");
            }
            catch (Exception ex)
            {
                // Handle any exception thrown due to worksheet protection
                Console.WriteLine("Error adding WordArt to a protected worksheet:");
                Console.WriteLine(ex.Message);
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            try
            {
                workbook.Save("WordArtProtectedSheet.xlsx");
                Console.WriteLine("Workbook saved as WordArtProtectedSheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook:");
                Console.WriteLine(ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("Unexpected error:");
            Console.WriteLine(ex.Message);
        }
    }
}
