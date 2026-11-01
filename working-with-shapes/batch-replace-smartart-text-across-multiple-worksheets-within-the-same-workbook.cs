// Title: Batch replace text in every SmartArt shape across all worksheets of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through each worksheet in a workbook and substitutes a given string in all SmartArt shapes with Aspose.Cells. | Show how to use dynamic typing to access SmartArt nodes and replace their paragraph portions in a .NET Excel automation scenario. | Provide an example that includes try‑catch blocks while updating SmartArt content across multiple sheets and then saves the workbook.
// Common Searches: Aspose.Cells replace text in SmartArt across multiple worksheets | C# batch update SmartArt node text in an Excel file | How to iterate worksheets and modify SmartArt shapes programmatically with Aspose.Cells | Dynamic access to SmartArt objects in Aspose.Cells .NET example | Error handling when processing SmartArt shapes in a workbook using Aspose.Cells
// Tags: batch replace smartart text aspocells | iterate worksheets smartart shapes c# | dynamic smartart node editing aspocells | excel workbook smartart text replacement | smartart processing error handling aspocells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an Excel workbook, iterates through every worksheet and each shape, identifies SmartArt objects, recursively replaces occurrences of a specified old string with a new string in all SmartArt node portions using dynamic typing, handles any processing errors, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            string oldText = "OldText";
            string newText = "NewText";

            // Iterate worksheets and shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        try
                        {
                            // Access SmartArt via dynamic to bypass compile‑time member check
                            dynamic smartArt = ((dynamic)shape).SmartArt;
                            if (smartArt != null)
                            {
                                ReplaceSmartArtText(smartArt, oldText, newText);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing SmartArt on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }
            }

            // Save modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Recursively replace text in all SmartArt nodes using dynamic typing
    static void ReplaceSmartArtText(dynamic smartArt, string oldText, string newText)
    {
        foreach (dynamic node in smartArt.AllNodes)
        {
            foreach (dynamic paragraph in node.Paragraphs)
            {
                foreach (dynamic portion in paragraph.Portions)
                {
                    if (portion.Text != null && portion.Text.Contains(oldText))
                    {
                        portion.Text = portion.Text.Replace(oldText, newText);
                    }
                }
            }
        }
    }
}
