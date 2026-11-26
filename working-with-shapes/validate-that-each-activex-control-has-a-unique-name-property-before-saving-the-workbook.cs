// Title: Validate that each ActiveX control in an Excel workbook has a unique name using Aspose.Cells for .NET before saving
// Common Searches: c# Aspose.Cells detect duplicate ActiveX control names before saving workbook | how to ensure unique ActiveX control names across worksheets with Aspose.Cells | validate shape name uniqueness for ActiveX controls in Excel using Aspose.Cells .NET | prevent Excel file save when ActiveX control names collide in Aspose.Cells | check for duplicate ActiveX control names in a workbook with Aspose.Cells C#
// Tags: duplicate ActiveX control name detection Aspose.Cells | shape name uniqueness validation Aspose.Cells | prevent workbook save on duplicate ActiveX names | HashSet duplicate name check C# | ActiveX control name validation Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ActiveXControlNameValidator
{
    // The example loads an Excel workbook with Aspose.Cells, iterates each worksheet and its shapes, and uses a case‑insensitive HashSet to track the Name property of shapes that host ActiveX controls. If a name is encountered more than once, the control is flagged as a duplicate and the workbook is not saved; otherwise the workbook is saved to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: ActiveXControlNameValidator <inputFilePath> [outputFilePath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : "ValidatedOutput.xlsx";

            if (!System.IO.File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                bool hasDuplicates = false;
                List<string> duplicateNames = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Use a HashSet to track names within the current worksheet
                    HashSet<string> namesInSheet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Only consider shapes that host an ActiveX control
                        if (shape.ActiveXControl != null)
                        {
                            string shapeName = shape.Name;

                            // If the shape has no name, treat it as an empty string (will be caught as duplicate if another unnamed shape exists)
                            if (string.IsNullOrEmpty(shapeName))
                            {
                                shapeName = string.Empty;
                            }

                            if (!namesInSheet.Add(shapeName))
                            {
                                // Duplicate found
                                hasDuplicates = true;
                                duplicateNames.Add($"Worksheet \"{sheet.Name}\": duplicate ActiveX control name \"{shapeName}\"");
                            }
                        }
                    }
                }

                if (hasDuplicates)
                {
                    Console.WriteLine("Validation failed: One or more ActiveX controls have duplicate names.");
                    foreach (string msg in duplicateNames)
                    {
                        Console.WriteLine(msg);
                    }
                    Console.WriteLine("The workbook will NOT be saved.");
                }
                else
                {
                    // No duplicates – safe to save
                    workbook.Save(outputPath);
                    Console.WriteLine($"Validation succeeded. Workbook saved to \"{outputPath}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
