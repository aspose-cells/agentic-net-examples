// Title: Unlock a Locked WordArt Watermark in Excel with Aspose.Cells (C#)
// Description: Load an Excel file, unprotect the worksheet if needed, iterate through all shapes, set Shape.IsLocked and TextBody.IsLockedText to false, and save the workbook so the WordArt watermark can be edited even when the sheet remains protected.
// Keywords: Aspose.Cells unlock WordArt | Excel shape IsLocked false | remove watermark protection C# | modify locked WordArt Aspose | unprotect worksheet Aspose.Cells | shape.TextBody.IsLockedText | C# Excel watermark editing
// Common Searches: how to unlock WordArt watermark in Excel using Aspose.Cells | set shape.IsLocked = false with Aspose.Cells C# | unlock text inside WordArt shape Aspose.Cells | unprotect worksheet and edit watermark programmatically | Aspose.Cells example for unlocking shapes
// Developer Intent: Programmatically remove lock restrictions from a WordArt watermark so it can be edited while the worksheet stays protected.
// Use Cases: Enable end‑users to change the text or position of a WordArt watermark in a protected template. | Batch‑process multiple workbooks to unlock watermarks before applying automated content updates. | Prepare reusable Excel templates where the watermark remains editable by downstream automation.
// AI Prompts: Write C# code with Aspose.Cells that unlocks all shapes and their text in a protected worksheet, then re‑applies protection with a password. | Show how to detect only WordArt shapes in a worksheet and unlock those while leaving other shapes unchanged. | Provide error‑handling for missing files, absent shapes, or incorrect worksheet passwords when unlocking a WordArt watermark.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsUnlockWordArtWatermark
{
    // Load an Excel file, unprotect the worksheet if needed, iterate through all shapes, set Shape.IsLocked and TextBody.IsLockedText to false, and save the workbook so the WordArt watermark can be edited even when the sheet remains protected.
    class Program
    {
        static void Main()
        {
            const string inputPath = "LockedWatermark.xlsx";
            const string outputPath = "UnlockedWatermark.xlsx";

            try
            {
                // Ensure the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains the locked WordArt watermark
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Unprotect the worksheet if it is protected
                if (worksheet.IsProtected)
                {
                    // Replace with actual password if required; empty string for none
                    worksheet.Unprotect("yourPassword");
                }

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Unlock the shape so it can be modified when the sheet is protected
                    shape.IsLocked = false;

                    // Unlock the text within the shape if it exists
                    if (shape.TextBody != null && shape.TextBody.TextAlignment != null)
                    {
                        shape.TextBody.TextAlignment.IsLockedText = false;
                    }
                }

                // Save the workbook with the unlocked watermark
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
