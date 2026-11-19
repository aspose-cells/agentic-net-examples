// Title: Validate that every Excel worksheet contains exactly one locked WordArt (TextEffect) watermark with Aspose.Cells for .NET
// Common Searches: how to ensure each sheet in an Excel workbook has a single locked WordArt watermark using Aspose.Cells C# | Aspose.Cells C# count locked TextEffect shapes per worksheet | report worksheets that are missing or have multiple locked WordArt watermarks in .NET | validate Excel file watermarks by checking shape.IsLocked property with Aspose.Cells | detect extra or absent WordArt watermarks in Excel sheets programmatically
// Tags: Aspose.Cells locked WordArt check | TextEffect shape count per sheet | Excel worksheet watermark audit | shape.IsLocked verification Aspose.Cells | single watermark validation per worksheet .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, counts shapes whose Type is TextEffect and IsLocked is true, and logs any sheet that does not have exactly one such locked WordArt watermark. It also handles missing files and runtime exceptions.
class WatermarkValidator
{
    // Validates that each worksheet contains exactly one locked WordArt watermark.
    // Reports any worksheets that do not meet this requirement.
    public static void ValidateWatermarks(string workbookPath)
    {
        // Ensure the file exists before attempting to load.
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int lockedWordArtCount = 0;

                // Examine each shape on the worksheet.
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify WordArt shapes (TextEffect) that are locked.
                    // Use string comparison to avoid direct dependency on ShapeType enum.
                    if (shape.IsLocked && shape.Type.ToString() == "TextEffect")
                    {
                        lockedWordArtCount++;
                    }
                }

                // Report discrepancy if the count is not exactly one.
                if (lockedWordArtCount != 1)
                {
                    Console.WriteLine(
                        $"Worksheet \"{sheet.Name}\" has {lockedWordArtCount} locked WordArt watermark(s). Expected exactly 1.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Example usage.
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as an argument.");
            return;
        }

        string filePath = args[0];
        ValidateWatermarks(filePath);
    }
}
