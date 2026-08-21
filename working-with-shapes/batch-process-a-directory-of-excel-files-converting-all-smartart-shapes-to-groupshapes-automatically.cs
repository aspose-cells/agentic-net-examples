// Title: Batch Convert SmartArt to GroupShape in Excel Workbooks using Aspose.Cells for .NET
// Description: A C# utility that scans a folder for .xlsx, .xlsm and .xlsb files, loads each workbook with LoadOptions.IgnoreUselessShapes, iterates every worksheet, transforms each SmartArt shape into a GroupShape via GetResultOfSmartArt, and saves the file with OoxmlSaveOptions.UpdateSmartArt to retain the conversion.
// Keywords: Aspose.Cells | C# | SmartArt conversion | GroupShape | batch processing Excel | LoadOptions.IgnoreUselessShapes | OoxmlSaveOptions.UpdateSmartArt | automate shape conversion | Excel workbook automation | folder scan
// Common Searches: convert all SmartArt to GroupShape in multiple Excel files | Aspose.Cells batch SmartArt conversion .NET | how to replace SmartArt with GroupShape using Aspose | ignore useless shapes for faster Excel processing | save workbook after SmartArt update Aspose.Cells
// Developer Intent: Automatically replace every SmartArt object with an equivalent GroupShape in all Excel workbooks located in a specified directory.
// Use Cases: Standardize diagrams across a portfolio of financial reports before distribution. | Prepare workbooks for platforms that do not support SmartArt by converting them to editable shapes. | Speed up large‑scale Excel processing by skipping irrelevant shapes while performing conversions.
// AI Prompts: Generate C# code that recursively processes subfolders and logs the number of SmartArt shapes converted per file. | Create a version of the batch converter that writes a CSV summary with workbook name, worksheet, and conversion count. | Show how to delete the original SmartArt after conversion and preserve its position and size in the new GroupShape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Utility;

// A C# utility that scans a folder for .xlsx, .xlsm and .xlsb files, loads each workbook with LoadOptions.IgnoreUselessShapes, iterates every worksheet, transforms each SmartArt shape into a GroupShape via GetResultOfSmartArt, and saves the file with OoxmlSaveOptions.UpdateSmartArt to retain the conversion.
class SmartArtBatchConverter
{
    static void Main(string[] args)
    {
        // Directory containing Excel files (change as needed)
        string sourceDirectory = @"C:\ExcelFiles";

        // Verify that the source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory not found: {sourceDirectory}");
            return;
        }

        // Process each Excel file in the directory (xlsx, xlsm, xlsb)
        foreach (string filePath in Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Skip files that are not Excel workbooks
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                continue;

            // Ensure the file actually exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load workbook with options (ignore useless shapes to speed up processing)
                LoadOptions loadOptions = new LoadOptions
                {
                    IgnoreUselessShapes = true
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Iterate through all worksheets and their shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;
                    // Use a copy of the collection count because GetResultOfSmartArt may add new shapes
                    int shapeCount = shapes.Count;
                    for (int i = 0; i < shapeCount; i++)
                    {
                        Shape shape = shapes[i];
                        if (shape.IsSmartArt)
                        {
                            // Convert SmartArt to a GroupShape
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            // Optional: adjust the position of the new group shape if needed
                            // groupShape.Left = shape.Left;
                            // groupShape.Top = shape.Top;
                        }
                    }
                }

                // Save the workbook, enabling UpdateSmartArt to persist the conversion
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(filePath, saveOptions);
                Console.WriteLine($"Processed and saved: {filePath}");
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
