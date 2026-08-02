// Title: Convert SmartArt to GroupShape in Excel using Aspose.Cells for .NET
// Description: Loads an XLSX file, scans every worksheet for SmartArt shapes, transforms each into a GroupShape via Shape.GetResultOfSmartArt, assigns a custom name, and saves the workbook with UpdateSmartArt enabled.
// Keywords: Aspose.Cells SmartArt conversion | Shape.GetResultOfSmartArt example | C# group shape from SmartArt | Excel workbook update SmartArt | iterate worksheet shapes Aspose
// Common Searches: How to replace SmartArt with GroupShape in Aspose.Cells | GetResultOfSmartArt C# code sample | Save Excel after converting SmartArt to group | Loop through shapes in all worksheets Aspose
// Developer Intent: Transform every SmartArt object in a workbook into an editable GroupShape and persist the change.
// Use Cases: Batch conversion of SmartArt for downstream styling or data extraction | Assigning identifiable names to converted groups for automated processing | Ensuring layout consistency by saving with UpdateSmartArt after conversion
// AI Prompts: Add comprehensive error handling for null results from GetResultOfSmartArt. | Show code to duplicate each converted GroupShape onto a summary worksheet. | Explain how to keep the original SmartArt layout while converting to a GroupShape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an XLSX file, scans every worksheet for SmartArt shapes, transforms each into a GroupShape via Shape.GetResultOfSmartArt, assigns a custom name, and saves the workbook with UpdateSmartArt enabled.
public class ConvertSmartArtToGroup
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load the workbook that contains SmartArt shapes
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through each shape in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Convert the SmartArt shape to a GroupShape
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // If conversion succeeded, further manipulate the group
                    if (groupShape != null)
                    {
                        // Assign a meaningful name to the new group
                        groupShape.Name = "ConvertedSmartArtGroup";
                    }
                }
            }
        }

        // Save the workbook with SmartArt update enabled so the conversion is persisted
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            UpdateSmartArt = true
        };
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Workbook saved successfully to {outputPath}");
    }
}
