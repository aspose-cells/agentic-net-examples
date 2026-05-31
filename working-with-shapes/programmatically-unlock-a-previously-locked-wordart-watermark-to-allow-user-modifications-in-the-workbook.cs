using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UnlockWordArtWatermark
{
    static void Main()
    {
        const string inputFile = "ProtectedWatermark.xlsx";
        const string outputFile = "UnlockedWatermark.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            // Load the workbook that contains the locked WordArt watermark
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Unprotect the worksheet if it is protected (no password assumed)
                if (worksheet.IsProtected)
                    worksheet.Unprotect();

                // Loop through all shapes on the worksheet
                for (int i = 0; i < worksheet.Shapes.Count; i++)
                {
                    Shape shape = worksheet.Shapes[i];

                    // Unlock the shape itself so it can be moved or resized
                    shape.IsLocked = false;

                    // If the shape contains a TextBody, unlock its text as well (if supported)
                    if (shape.TextBody != null)
                    {
                        // Aspose.Cells does not expose an IsLocked property for TextBody,
                        // so we rely on unlocking the shape itself.
                        // Additional text unlocking logic can be added here if needed.
                    }
                }
            }

            // Save the workbook with the unlocked watermark
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}