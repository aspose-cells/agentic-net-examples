// Title: How to replace a placeholder tag in a TextBox shape of an Excel workbook with a dynamic title using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace placeholder text in a textbox shape | Find a TextBox shape by name and change its text with Aspose.Cells | Replace <TAG_1> with a dynamic title in an Excel workbook using Aspose.Cells .NET | Iterate through worksheet shapes to modify textbox content in C# | Aspose.Cells update shape text based on a placeholder tag
// Tags: Aspose.Cells modify TextBox shape content | search shape by name in Excel workbook | insert dynamic title into Excel textbox | C# iterate worksheet shapes with Aspose.Cells | update shape text using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, iterates over each worksheet's shapes to locate a TextBox named "HeaderBox", replaces the placeholder "<TAG_1>" with a specified dynamic title, and saves the updated workbook, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetShapeName = "HeaderBox";
        const string placeholderTag = "<TAG_1>";
        const string dynamicTitle = "My Dynamic Title";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify the shape by its name (assumed to be a TextBox)
                    if (shape.Name == targetShapeName)
                    {
                        // Retrieve current text, replace the placeholder, and assign back
                        string currentText = shape.Text;
                        if (!string.IsNullOrEmpty(currentText) && currentText.Contains(placeholderTag))
                        {
                            shape.Text = currentText.Replace(placeholderTag, dynamicTitle);
                        }
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
