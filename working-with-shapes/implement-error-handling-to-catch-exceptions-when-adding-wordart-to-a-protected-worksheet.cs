// Title: C# – Add WordArt to a Protected Worksheet with Aspose.Cells and Exception Handling
// Description: Demonstrates how to protect a worksheet for objects, attempt to insert WordArt using ShapeCollection.AddWordArt, and gracefully catch the CellsException (ExceptionType.Shape) when protection blocks the operation. Includes a fallback catch for any other errors and saves the workbook.
// Keywords: Aspose.Cells C# | AddWordArt protected worksheet | CellsException Shape handling | worksheet protection objects | exception handling Aspose.Cells | try‑catch shape addition | WordArt Aspose.Cells .NET
// Common Searches: add wordart to a protected sheet aspose.cells | cellsexception shape when adding shape to protected worksheet | how to catch shape errors in aspose.cells c# | protect worksheet objects and insert drawing objects aspose | asp.net example wordart exception handling
// Developer Intent: Insert WordArt into a worksheet that is locked for objects and handle the resulting shape‑related exception.
// Use Cases: Add decorative WordArt to a report while the sheet is object‑protected, notifying the user if the operation fails. | Log detailed CellsException information for diagnostics when protection prevents shape creation. | Automatically switch to a plain text box if WordArt cannot be added due to object protection.
// AI Prompts: Write C# code that adds WordArt to a worksheet protected with ProtectionType.Objects using Aspose.Cells and includes try‑catch for CellsException with ExceptionType.Shape. | Explain why Aspose.Cells throws a CellsException with code Shape when a shape is added to a protected sheet and outline best practices for handling it. | Suggest alternative visual elements that can be added to a protected worksheet without triggering shape protection errors in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect a worksheet for objects, attempt to insert WordArt using ShapeCollection.AddWordArt, and gracefully catch the CellsException (ExceptionType.Shape) when protection blocks the operation. Includes a fallback catch for any other errors and saves the workbook.
    public class AddWordArtToProtectedWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Protect the worksheet so that drawing objects cannot be edited/added
                worksheet.Protect(ProtectionType.Objects);

                // Get the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add WordArt with specified parameters
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Protected WordArt",
                    2,      // topRow
                    0,      // top (pixel offset)
                    2,      // leftColumn
                    0,      // left (pixel offset)
                    100,    // height (pixels)
                    300     // width (pixels)
                );

                Console.WriteLine("WordArt added successfully.");

                // Save the workbook
                workbook.Save("WordArtProtected.xlsx");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Shape)
            {
                // Handle the specific exception thrown when adding a shape to a protected sheet
                Console.WriteLine("Failed to add WordArt: " + ex.Message);
                Console.WriteLine("Exception Type: " + ex.Code);
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddWordArtToProtectedWorksheet.Run();
        }
    }
}
