// Title: Unlock a Signature Shape in Excel with Aspose.Cells for .NET – Enable Move & Resize
// Description: Loads an Excel workbook, finds the shape named "Signature" on the first worksheet, asks the user for confirmation, clears the shape's lock and its move/resize restrictions, optionally re‑protects the sheet, and saves the file as an unlocked version.
// Keywords: Aspose.Cells unlock shape | Excel signature shape move | resize locked shape .NET | shape IsLocked false | ShapeLockType Move Resize | unlock shape after confirmation | protected worksheet shape edit
// Common Searches: how to unlock a signature shape in Excel using Aspose.Cells | enable moving and resizing of a locked shape in .NET | Aspose.Cells unlock shape while worksheet is protected | C# code to change shape lock properties in Excel | unlock Excel shape after user prompt
// Developer Intent: Remove the lock on the "Signature" shape so it can be moved or resized after user approval.
// Use Cases: Prompt users before unlocking a signature shape on a protected sheet. | Programmatically adjust the size or position of a locked signature without removing protection. | Batch‑unlock multiple named shapes to apply layout changes in an automated workflow.
// AI Prompts: Generate C# code with Aspose.Cells that locates a shape called "Signature", asks for user confirmation, and unlocks its move and resize properties. | Show how to unlock a shape, keep the worksheet protected, and save the workbook using Aspose.Cells for .NET. | Create a reusable method that takes a workbook path and shape name, unlocks the shape for editing, and returns the updated workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureUnlock
{
    // Loads an Excel workbook, finds the shape named "Signature" on the first worksheet, asks the user for confirmation, clears the shape's lock and its move/resize restrictions, optionally re‑protects the sheet, and saves the file as an unlocked version.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "SignedDocument.xlsx";
                const string outputPath = "SignedDocument_Unlocked.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Please place the workbook in the application directory.");
                    return;
                }

                // Load the workbook that contains the signature shape
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0]; // assume the shape is on the first sheet

                // Locate the shape named "Signature"
                Shape signatureShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name.Equals("Signature", StringComparison.OrdinalIgnoreCase))
                    {
                        signatureShape = shape;
                        break;
                    }
                }

                if (signatureShape == null)
                {
                    Console.WriteLine("Signature shape not found.");
                    return;
                }

                // Ask user for confirmation before unlocking
                Console.Write("Do you want to unlock the signature shape for editing? (y/n): ");
                string answer = Console.ReadLine();
                if (!answer.Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Operation cancelled by user.");
                    return;
                }

                // Unlock the shape so it can be moved and resized even when the worksheet is protected
                signatureShape.IsLocked = false; // general lock
                signatureShape.SetLockedProperty(ShapeLockType.Move, false);   // allow moving
                signatureShape.SetLockedProperty(ShapeLockType.Resize, false); // allow resizing

                // Optionally, protect the worksheet again while keeping the shape unlocked
                // worksheet.Protect(ProtectionType.All);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Signature shape unlocked and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
