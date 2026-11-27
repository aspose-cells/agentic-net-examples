// Title: C# script to unlock Excel shapes that have a specific AlternativeText tag using Aspose.Cells
// Common Searches: how to unlock Excel shapes with a given AlternativeText using Aspose.Cells C# | C# iterate all shapes in a worksheet and set IsLocked to false with Aspose.Cells | temporarily unprotect a worksheet, modify shape lock state, then reapply protection Aspose.Cells | search for shapes by custom tag and unlock them in an .xlsx file using Aspose.Cells
// Tags: unlock shape IsLocked Aspose.Cells | filter shapes by AlternativeText C# | temporarily unprotect worksheet Aspose.Cells | reapply object protection Aspose.Cells | iterate worksheet shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockShapesByTag
{
    // The program loads an Excel workbook, temporarily removes worksheet protection, iterates through each shape, and unlocks any shape whose AlternativeText matches the supplied tag by setting IsLocked to false. After processing, it restores object‑level protection and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: UnlockShapesByTag <input.xlsx> <output.xlsx> <tagValue>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string targetTag = args[2];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool wasProtected = sheet.IsProtected;

                    // If the sheet is protected, temporarily unprotect it
                    if (wasProtected)
                    {
                        sheet.Unprotect();
                    }

                    // Iterate through all shapes in the worksheet
                    ShapeCollection shapes = sheet.Shapes;
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Use AlternativeText as the custom tag storage.
                        if (string.Equals(shape.AlternativeText, targetTag, StringComparison.OrdinalIgnoreCase))
                        {
                            // Unlock the shape for editing
                            shape.IsLocked = false;
                            Console.WriteLine($"Unlocked shape '{shape.Name}' (Index {i}) on sheet '{sheet.Name}'.");
                        }
                    }

                    // Re‑apply protection if it was originally set (protect objects only)
                    if (wasProtected)
                    {
                        sheet.Protect(ProtectionType.Objects);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
