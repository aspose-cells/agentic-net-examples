// Title: How to catch exceptions when adding WordArt (TextEffect) to a protected worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add WordArt to a protected sheet with try-catch | handling exception for AddTextEffect on a protected worksheet in Aspose.Cells | Aspose.Cells .NET prevent crash when inserting TextEffect into a locked worksheet | example of protecting worksheet and safely adding WordArt using Aspose.Cells | catching shape insertion errors in a protected Excel sheet with Aspose.Cells
// Tags: AddTextEffect exception handling Aspose.Cells | protected worksheet shape insertion C# | WordArt shape addition error handling .NET | Aspose.Cells worksheet protection API | try-catch pattern for Excel shape operations

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, protects the first worksheet with all protection types, and attempts to add a WordArt (TextEffect) shape via AddTextEffect. Because the sheet is protected, an exception is thrown and caught in a try‑catch block, logging the error. The workbook is then saved inside another try‑catch to handle any saving issues.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet (all protection types)
        sheet.Protect(ProtectionType.All);

        try
        {
            // Attempt to add WordArt (TextEffect) to the protected worksheet.
            // This will throw an exception because the sheet is protected.
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // upper left row, upper left column, lower right row, lower right column,
            // height, width.
            Shape shape = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset effect
                "Hello World",                  // text
                "Arial",                        // font name
                36,                             // font size
                false,                          // bold
                false,                          // italic
                5,                              // upper left row
                2,                              // upper left column
                10,                             // lower right row
                5,                              // lower right column
                200,                            // height
                400);                           // width

            // shape variable can be used if needed
        }
        catch (Exception ex)
        {
            // Handle the exception (e.g., log or display the error message)
            Console.WriteLine("Error adding WordArt to protected worksheet: " + ex.Message);
        }

        try
        {
            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving workbook: " + ex.Message);
        }
    }
}
