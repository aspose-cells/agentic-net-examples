// Title: Clone a SmartArt shape (placeholder) and compare GroupShape child counts with Aspose.Cells for .NET
// Description: Loads a workbook, finds the first SmartArt shape, converts it to a GroupShape using GetResultOfSmartArt to obtain the child shape count, adds a rectangle as a placeholder for a cloned SmartArt, attempts conversion of the placeholder to a GroupShape, compares the original and cloned child counts, and saves the result.
// Keywords: Aspose.Cells | SmartArt | GroupShape | GetResultOfSmartArt | child shape count | clone SmartArt | C# | .NET | shape placeholder | AddShape | compare shapes
// Common Searches: Aspose.Cells clone SmartArt | Get child shapes from SmartArt Aspose | Convert SmartArt to GroupShape .NET | Why GetResultOfSmartArt returns null | Duplicate SmartArt in Aspose.Cells | Count grouped shapes in SmartArt | AddShape placeholder SmartArt Aspose
// Developer Intent: Attempt to duplicate a SmartArt shape, convert both the original and the placeholder to GroupShape objects, and evaluate whether their child shape counts match.
// Use Cases: Validate that a copied SmartArt retains the same internal hierarchy by comparing child counts. | Generate diagnostics for the limitation of cloning SmartArt with a rectangle placeholder. | Create a report showing original SmartArt child count versus placeholder conversion result. | Write unit tests for SmartArt conversion and cloning behavior in Aspose.Cells.
// AI Prompts: Write C# code that clones a SmartArt shape by copying its properties and compares child counts of the original and cloned GroupShape using Aspose.Cells. | Explain why GetResultOfSmartArt returns null for shapes created with AddShape and suggest a proper method to duplicate SmartArt in Aspose.Cells. | Provide a NUnit test that loads a workbook, converts a SmartArt to GroupShape, asserts the expected child count, and verifies the outcome when attempting to clone the SmartArt.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtCloneDemo
{
    // Loads a workbook, finds the first SmartArt shape, converts it to a GroupShape using GetResultOfSmartArt to obtain the child shape count, adds a rectangle as a placeholder for a cloned SmartArt, attempts conversion of the placeholder to a GroupShape, compares the original and cloned child counts, and saves the result.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string inputFile = "SmartArtSample.xlsx";
                const string outputFile = "SmartArtCloneResult.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found. Please ensure the file exists in the application directory.");
                    return;
                }

                // Load the workbook that contains a SmartArt shape
                Workbook workbook = new Workbook(inputFile);
                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // Locate the first SmartArt shape in the worksheet
                SmartArtShape smartArt = null;
                foreach (Shape s in shapes)
                {
                    if (s is SmartArtShape sa)
                    {
                        smartArt = sa;
                        break;
                    }
                }

                if (smartArt == null)
                {
                    Console.WriteLine("No SmartArt shape found in the worksheet.");
                    return;
                }

                // Convert the original SmartArt to a GroupShape
                GroupShape originalGroup = smartArt.GetResultOfSmartArt();

                // Get the number of child shapes inside the original group
                int originalChildCount = originalGroup?.GetGroupedShapes()?.Length ?? 0;
                Console.WriteLine($"Original GroupShape contains {originalChildCount} child shapes.");

                // ----- Clone the SmartArt shape (illustrative) -----
                // Since Aspose.Cells does not provide a direct Clone method for SmartArt,
                // we add a rectangle with the same bounds as a placeholder.
                Shape clonedShape = shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    smartArt.UpperLeftRow,
                    smartArt.UpperLeftColumn,
                    smartArt.UpperLeftRow,
                    smartArt.UpperLeftColumn,
                    smartArt.Width,
                    smartArt.Height);

                // Attempt to convert the cloned shape to a GroupShape.
                // This will return null because the placeholder is not a true SmartArt.
                GroupShape clonedGroup = clonedShape.GetResultOfSmartArt();

                if (clonedGroup != null)
                {
                    int clonedChildCount = clonedGroup.GetGroupedShapes().Length;
                    Console.WriteLine($"Cloned GroupShape contains {clonedChildCount} child shapes.");

                    // Compare child counts
                    if (originalChildCount == clonedChildCount)
                    {
                        Console.WriteLine("The original and cloned SmartArt have the same number of child shapes.");
                    }
                    else
                    {
                        Console.WriteLine("The original and cloned SmartArt have different numbers of child shapes.");
                    }
                }
                else
                {
                    Console.WriteLine("Cloned shape could not be converted to GroupShape (not a SmartArt).");
                    Console.WriteLine($"Original child count: {originalChildCount}, cloned child count: N/A");
                }

                // Save the workbook with the added cloned shape
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
