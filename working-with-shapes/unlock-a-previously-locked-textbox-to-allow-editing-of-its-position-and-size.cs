// Title: Unlock a locked textbox shape and disable its aspect‑ratio lock in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unlock textbox shape to move and resize | How to disable aspect ratio lock for Excel shapes with Aspose.Cells .NET | Unlock locked shape properties in Aspose.Cells workbook via reflection | Change Locked and LockAspectRatio of a textbox using Aspose.Cells C# example
// Tags: unlock textbox shape Aspose.Cells | disable aspect ratio lock Aspose.Cells | modify Locked property shape .NET | reflection set shape properties Aspose.Cells | edit textbox position size Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, iterates through all shapes on the first worksheet, and uses reflection to set the Locked and LockAspectRatio properties of each shape to false, unlocking the textbox and allowing its position and size to be edited before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Attempt to unlock the shape and disable aspect‑ratio locking using reflection.
                // This avoids compile‑time errors if the properties are not present in the used version.
                var lockedProp = shape.GetType().GetProperty("Locked");
                if (lockedProp != null && lockedProp.CanWrite)
                {
                    lockedProp.SetValue(shape, false);
                }

                var lockAspectProp = shape.GetType().GetProperty("LockAspectRatio");
                if (lockAspectProp != null && lockAspectProp.CanWrite)
                {
                    lockAspectProp.SetValue(shape, false);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
