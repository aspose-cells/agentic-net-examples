// Title: Unlock WordArt Watermark Shapes in Excel with Aspose.Cells (.NET)
// Description: Loads a password‑protected workbook, removes workbook and worksheet protection, iterates all shapes, clears the IsLocked flag and the IsLockedText property for WordArt, then saves the unlocked file.
// Keywords: Aspose.Cells unlock WordArt | Excel shape unlock .NET | remove locked watermark Aspose | unprotect worksheet shapes C# | edit WordArt text programmatically | IsLockedText false Aspose.Cells | batch unlock Excel watermarks
// Common Searches: how to unlock WordArt watermark in Excel using Aspose.Cells | C# code to make locked shapes editable after workbook protection | remove shape lock from protected worksheet Aspose.Cells | unlock text editing for WordArt in a protected Excel file | Aspose.Cells unlock watermark without losing other protections
// Developer Intent: Programmatically remove the lock on WordArt watermark shapes so they can be edited after the workbook and worksheet are unprotected.
// Use Cases: Enable end‑users to modify the text of a pre‑locked WordArt watermark in a template. | Automate unlocking of shape objects across many workbooks before applying bulk updates. | Prepare a protected Excel template, then programmatically unlock only the watermark while keeping other protections intact. | Integrate into a CI pipeline to ensure watermarks are editable before publishing reports.
// AI Prompts: Write a C# method using Aspose.Cells that accepts a file path and password, unlocks all shapes, and clears IsLockedText for WordArt. | Explain the steps to safely unprotect a workbook and worksheet, then iterate shapes to change lock properties. | Provide error‑handling code for missing files or incorrect passwords when unlocking WordArt watermarks. | Show how to detect non‑WordArt shapes and skip them while unlocking only WordArt objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockWordArtWatermarkApp
{
    // Loads a password‑protected workbook, removes workbook and worksheet protection, iterates all shapes, clears the IsLocked flag and the IsLockedText property for WordArt, then saves the unlocked file.
    class UnlockWordArtWatermark
    {
        static void Main()
        {
            const string inputPath = "Watermarked.xlsx";
            const string outputPath = "Watermarked_Unlocked.xlsx";
            const string password = "password";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing the locked WordArt watermark
                Workbook workbook = new Workbook(inputPath);

                // Unprotect the workbook if it is password‑protected
                workbook.Unprotect(password);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Unprotect(password);

                // Unlock all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    shape.IsLocked = false;

                    // Unlock text editing for WordArt shapes
                    if (shape.TextBody != null && shape.TextBody.TextAlignment != null)
                    {
                        shape.TextBody.TextAlignment.IsLockedText = false;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
