// Title: C# – Handle Exceptions When Adding WordArt to a Protected Worksheet with Aspose.Cells
// Description: Demonstrates how to protect a worksheet, disable editing of drawing objects, and safely attempt to add WordArt using Aspose.Cells. The sample wraps the AddWordArt call in try‑catch blocks for CellsException and generic Exception, logs the error details, and ensures the workbook is saved whether the shape is added or not.
// Keywords: Aspose.Cells | C# | WordArt | protected worksheet | shape insertion error | CellsException | exception handling | save workbook after failure
// Common Searches: Aspose.Cells add WordArt to protected sheet C# | catch CellsException when inserting shape | error adding drawing objects to a protected worksheet | how to protect worksheet and still add WordArt | sample code for exception handling in Aspose.Cells
// Developer Intent: Insert WordArt into a worksheet while gracefully handling protection‑related errors.
// Use Cases: Protect a sheet, lock drawing objects, and attempt to add WordArt without crashing the app. | Capture and log the specific CellsException code when shape insertion is blocked by protection. | Guarantee that the workbook is saved either with the new WordArt or as a protected file without it.
// AI Prompts: Write C# code that adds WordArt to an Aspose.Cells worksheet and includes try‑catch blocks for CellsException and generic Exception, ensuring the file is saved in all cases. | Explain how worksheet protection settings influence shape operations in Aspose.Cells and how to retrieve the exception code when AddWordArt fails. | Refactor the example to log errors to a file instead of the console while still handling protected‑worksheet scenarios.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect a worksheet, disable editing of drawing objects, and safely attempt to add WordArt using Aspose.Cells. The sample wraps the AddWordArt call in try‑catch blocks for CellsException and generic Exception, logs the error details, and ensures the workbook is saved whether the shape is added or not.
    public class AddWordArtWithProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Disallow editing of drawing objects on a protected worksheet
                worksheet.Protection.AllowEditingObject = false;

                // Protect the worksheet with a password
                worksheet.Protect(ProtectionType.All, "password123", null);

                // Attempt to add WordArt to the protected worksheet
                try
                {
                    // Get the shape collection of the worksheet
                    ShapeCollection shapes = worksheet.Shapes;

                    // This call will throw an exception because the worksheet is protected
                    Shape wordArt = shapes.AddWordArt(
                        PresetWordArtStyle.WordArtStyle2, // preset style
                        "Protected WordArt",              // text
                        2,   // top row index
                        0,   // top offset (pixels)
                        2,   // left column index
                        0,   // left offset (pixels)
                        100, // height (pixels)
                        400  // width (pixels)
                    );

                    // If no exception, save the workbook with WordArt
                    workbook.Save("WordArtAdded.xlsx");
                }
                catch (CellsException ex) // catches Aspose.Cells specific exceptions
                {
                    Console.WriteLine("Failed to add WordArt to a protected worksheet.");
                    Console.WriteLine($"Exception Message: {ex.Message}");
                    Console.WriteLine($"Exception Type Code: {ex.Code}");
                }
                catch (Exception ex) // catches any other unexpected exceptions
                {
                    Console.WriteLine("An unexpected error occurred while adding WordArt.");
                    Console.WriteLine($"Exception Message: {ex.Message}");
                }

                // Save the workbook (even if WordArt was not added)
                workbook.Save("ProtectedWorksheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during workbook processing.");
                Console.WriteLine($"Exception Message: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddWordArtWithProtectionDemo.Run();
        }
    }
}
