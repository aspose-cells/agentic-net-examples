// Title: Apply Text Rotation to a Merged Named Range with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines the range B2:D4, converts it to a UnionRange, names it "MyNamedRange", merges the cells, inserts "Rotated Text", builds a Style with a 45° RotationAngle, enables the rotation flag, applies the style to the merged named range, and saves the file as MergedNamedRangeWithRotation.xlsx.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET | rotate text in merged cells | named range styling | UnionRange Aspose | StyleFlag rotation | Excel text orientation | apply style to range | merged cells rotation | Excel automation Aspose
// Common Searches: rotate text in a merged named range Aspose.Cells | apply style with rotation to UnionRange C# | set text orientation for merged cells using Aspose.Cells | how to name and merge a range then rotate text | Aspose.Cells example for text rotation in merged cells
// Developer Intent: Create a merged named range and apply a 45‑degree text rotation style using Aspose.Cells for .NET.
// Use Cases: Designing a report header that spans columns with diagonal text for compact labeling. | Building a template where merged title cells need rotated labels to group sub‑columns. | Automating workbook generation that requires rotated annotations inside merged cells for visual emphasis.
// AI Prompts: Show C# code to change the rotation angle of a styled merged named range in Aspose.Cells. | Provide an Aspose.Cells example that applies different text orientations to multiple merged named ranges. | Explain how StyleFlag can be used to apply only rotation without altering other formatting in a merged range.

using System;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example creates a workbook, defines the range B2:D4, converts it to a UnionRange, names it "MyNamedRange", merges the cells, inserts "Rotated Text", builds a Style with a 45° RotationAngle, enables the rotation flag, applies the style to the merged named range, and saves the file as MergedNamedRangeWithRotation.xlsx.
    public class ApplyRotationToMergedNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the address of the range to be named and merged
                string rangeAddress = "B2:D4";

                // Create a Range object for the specified address
                AsposeRange baseRange = worksheet.Cells.CreateRange(rangeAddress);

                // Convert the Range to a UnionRange (required for naming and merging)
                UnionRange unionRange = baseRange.UnionRanges(new AsposeRange[] { baseRange });

                // Assign a name to the UnionRange (named range)
                unionRange.Name = "MyNamedRange";

                // Merge the cells in the UnionRange
                unionRange.Merge();

                // Put a sample value into the merged cell (upper‑left cell of the range)
                unionRange.PutValue("Rotated Text", true, true);

                // Create a style with the desired text rotation
                Style rotationStyle = workbook.CreateStyle();
                rotationStyle.RotationAngle = 45; // Rotate text 45 degrees

                // Enable the rotation flag so the rotation is applied
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.Rotation = true;

                // Apply the style with rotation to the merged named range
                unionRange.ApplyStyle(rotationStyle, styleFlag);

                // Save the workbook
                string outputPath = "MergedNamedRangeWithRotation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyRotationToMergedNamedRange.Run();
        }
    }
}
