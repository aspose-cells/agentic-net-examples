// Title: Check each Excel worksheet for exactly one locked WordArt watermark with Aspose.Cells (.NET)
// Description: This C# example loads a workbook, scans every worksheet’s Shapes collection, counts shapes where IsWordArt and IsLocked are true, and reports sheets that have zero or more than one such watermark, ensuring consistent protection across the file.
// Keywords: Aspose.Cells | C# WordArt watermark | locked WordArt | Excel shape validation | worksheet watermark check | IsWordArt | IsLocked | Excel compliance | C# Excel automation | Aspose.Cells .NET | GitHub example
// Common Searches: How to verify a single locked WordArt watermark per sheet using Aspose.Cells | Aspose.Cells count WordArt shapes in C# | Validate Excel watermarks with Aspose.Cells .NET | C# code to detect missing WordArt watermark in workbook | GitHub Aspose.Cells watermark validation example
// Developer Intent: Confirm that every worksheet contains exactly one locked WordArt watermark and list any sheets that violate this rule.
// Use Cases: Quality‑control for automatically generated reports that must carry a protected watermark on each sheet. | Compliance audit of financial workbooks requiring a single locked WordArt watermark. | Pre‑publish script that flags or fixes worksheets with missing or duplicate watermarks. | Integration into CI pipelines to enforce watermark standards before release.
// AI Prompts: Generate C# code that adds a locked WordArt watermark to worksheets missing one using Aspose.Cells. | Rewrite the validation loop with LINQ to count locked WordArt shapes per worksheet. | Create a method returning a collection of worksheet names where the locked WordArt count is not exactly one. | Provide a PowerShell script that runs the compiled validator across multiple files. | Show how to log validation results to a JSON file for downstream processing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example loads a workbook, scans every worksheet’s Shapes collection, counts shapes where IsWordArt and IsLocked are true, and reports sheets that have zero or more than one such watermark, ensuring consistent protection across the file.
class WatermarkValidator
{
    static void Main(string[] args)
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            bool allWorksheetsValid = true;

            // Iterate through each worksheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                int lockedWordArtCount = 0;

                // Examine all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // WordArt can be identified via IsWordArt property
                    // Check if the shape is a WordArt and is locked
                    if (shape.IsWordArt && shape.IsLocked)
                    {
                        lockedWordArtCount++;
                    }
                }

                // Validate count
                if (lockedWordArtCount != 1)
                {
                    allWorksheetsValid = false;
                    Console.WriteLine($"Worksheet '{sheet.Name}' contains {lockedWordArtCount} locked WordArt watermarks (expected exactly 1).");
                }
            }

            if (allWorksheetsValid)
            {
                Console.WriteLine("All worksheets contain exactly one locked WordArt watermark.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
