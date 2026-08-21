// Title: Unit Test for Verifying SmartArt Text Replacement with Aspose.Cells (.NET)
// Description: Loads a workbook containing SmartArt, replaces every inner shape's text, saves with UpdateSmartArt, reloads the file, and asserts that the new text persists. Demonstrates a reliable test pattern for Aspose.Cells SmartArt modifications in C#.
// Keywords: Aspose.Cells | SmartArt | text replacement | unit test | C# | .NET | Excel automation | OoxmlSaveOptions | UpdateSmartArt | MSTest | xUnit | NUnit | automated testing
// Common Searches: Aspose.Cells unit test SmartArt text replacement | C# verify SmartArt changes after save | How to test SmartArt updates with Aspose.Cells | UpdateSmartArt option unit test example | Automated test for Excel SmartArt using Aspose
// Developer Intent: Ensure that modifying SmartArt node text via Aspose.Cells is correctly written to the workbook and remains after the file is saved and reopened.
// Use Cases: Continuous‑integration test that validates SmartArt label updates in generated reports. | Regression suite for a feature that customizes SmartArt captions before distribution. | Quality‑gate check confirming that the UpdateSmartArt flag preserves text changes across library versions.
// AI Prompts: Create an MSTest method that calls the SmartArt replacement routine and asserts the new text exists after saving the workbook. | Write an xUnit test case for SmartArt text replacement with temporary file handling and cleanup using Aspose.Cells. | Generate a NUnit test that verifies the UpdateSmartArt option persists SmartArt modifications and provides detailed failure messages.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // Loads a workbook containing SmartArt, replaces every inner shape's text, saves with UpdateSmartArt, reloads the file, and asserts that the new text persists. Demonstrates a reliable test pattern for Aspose.Cells SmartArt modifications in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                ReplaceSmartArtText();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void ReplaceSmartArtText()
        {
            const string templatePath = "SmartArtTemplate.xlsx";
            const string outputPath = "SmartArtResult.xlsx";
            const string newText = "Replaced";

            // Ensure the template file exists before loading.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that already contains a SmartArt shape.
            Workbook workbook = new Workbook(templatePath);

            // Iterate through worksheets and shapes, replace SmartArt text.
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt to a group of shapes.
                        GroupShape group = shape.GetResultOfSmartArt();

                        // Replace the text of each individual shape inside the SmartArt.
                        foreach (Shape smartArtShape in group.GetGroupedShapes())
                        {
                            smartArtShape.Text = newText;
                        }
                    }
                }
            }

            // Save the workbook with UpdateSmartArt enabled so that changes persist.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateSmartArt = true
            };
            workbook.Save(outputPath, saveOptions);

            // Verify that the SmartArt text was updated.
            if (!File.Exists(outputPath))
            {
                Console.WriteLine($"Failed to save output file: {outputPath}");
                return;
            }

            Workbook savedWorkbook = new Workbook(outputPath);
            bool replacementFound = false;

            foreach (Worksheet worksheet in savedWorkbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        GroupShape group = shape.GetResultOfSmartArt();
                        foreach (Shape smartArtShape in group.GetGroupedShapes())
                        {
                            if (smartArtShape.Text == newText)
                            {
                                replacementFound = true;
                                break;
                            }
                        }
                    }
                    if (replacementFound) break;
                }
                if (replacementFound) break;
            }

            Console.WriteLine(replacementFound
                ? "SmartArt text replacement was applied correctly."
                : "SmartArt text replacement was not applied.");
        }
    }
}
