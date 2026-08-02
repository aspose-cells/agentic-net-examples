// Title: C# Unit Test for Verifying SmartArt Text Replacement with Aspose.Cells
// Description: Loads a workbook containing SmartArt, converts each SmartArt shape to a GroupShape via GetResultOfSmartArt, replaces the text of every grouped element, saves the file with OoxmlSaveOptions.UpdateSmartArt enabled, reloads the workbook, and asserts that all SmartArt parts retain the new text.
// Keywords: Aspose.Cells | SmartArt | C# unit test | GetResultOfSmartArt | UpdateSmartArt | grouped shapes | text replacement | Excel automation | test automation | Workbook verification
// Common Searches: Aspose.Cells unit test SmartArt text | How to verify SmartArt changes after saving | C# test for UpdateSmartArt option | SmartArt text replacement example .NET | Validate grouped shape text in Excel with Aspose
// Developer Intent: Confirm that modifying SmartArt text persists after saving the workbook with the UpdateSmartArt flag.
// Use Cases: Automated regression testing for SmartArt modifications in reporting pipelines. | Continuous‑integration checks that custom SmartArt labels are correctly applied. | Ensuring downstream processes receive Excel files with updated SmartArt content.
// AI Prompts: Generate an MSTest method that loads a workbook, replaces all SmartArt text with "ReplacedText", saves using OoxmlSaveOptions.UpdateSmartArt = true, reloads the file, and asserts the replacement succeeded. | Write a NUnit test in C# that verifies SmartArt text persistence after serialization with Aspose.Cells. | Provide an xUnit example that mocks file I/O and checks that GetResultOfSmartArt returns grouped shapes whose Text property equals the expected value.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsTests
{
    // Loads a workbook containing SmartArt, converts each SmartArt shape to a GroupShape via GetResultOfSmartArt, replaces the text of every grouped element, saves the file with OoxmlSaveOptions.UpdateSmartArt enabled, reloads the workbook, and asserts that all SmartArt parts retain the new text.
    public class Program
    {
        private const string TemplateFile = "SmartArtTemplate.xlsx";
        private const string OutputFile = "SmartArtOutput.xlsx";

        public static void Main()
        {
            try
            {
                // Ensure the template file exists
                if (!File.Exists(TemplateFile))
                {
                    Console.WriteLine($"Template file \"{TemplateFile}\" not found.");
                    return;
                }

                // Delete any previous output file
                if (File.Exists(OutputFile))
                {
                    File.Delete(OutputFile);
                }

                // Load the workbook containing SmartArt shapes
                Workbook workbook = new Workbook(TemplateFile);

                // Replace text inside each SmartArt shape
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (shape.IsSmartArt)
                        {
                            // Convert the SmartArt to its grouped shape representation
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            // Iterate over the grouped shapes that represent individual SmartArt elements
                            foreach (Shape smartArtPart in groupShape.GetGroupedShapes())
                            {
                                // Replace the existing text with a known value
                                smartArtPart.Text = "ReplacedText";
                            }
                        }
                    }
                }

                // Save the workbook with UpdateSmartArt enabled so that the changes are persisted
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(OutputFile, saveOptions);

                // Reload the saved workbook to verify the changes
                Workbook reloaded = new Workbook(OutputFile);
                bool allTextsReplaced = true;

                foreach (Worksheet worksheet in reloaded.Worksheets)
                {
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (shape.IsSmartArt)
                        {
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            foreach (Shape smartArtPart in groupShape.GetGroupedShapes())
                            {
                                if (smartArtPart.Text != "ReplacedText")
                                {
                                    allTextsReplaced = false;
                                    break;
                                }
                            }
                        }

                        if (!allTextsReplaced) break;
                    }

                    if (!allTextsReplaced) break;
                }

                // Output verification result
                if (allTextsReplaced)
                {
                    Console.WriteLine("SmartArt text was replaced correctly after saving with UpdateSmartArt enabled.");
                }
                else
                {
                    Console.WriteLine("SmartArt text was NOT replaced correctly.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
