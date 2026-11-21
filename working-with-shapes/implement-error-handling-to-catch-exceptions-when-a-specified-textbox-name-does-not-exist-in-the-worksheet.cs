// Title: Catch an exception when a TextBox shape name does not exist in an Aspose.Cells worksheet (C#)
// Common Searches: C# Aspose.Cells handle missing TextBox shape exception | How to verify a TextBox shape exists before accessing it in Aspose.Cells | Retrieve TextBox by name in Aspose.Cells and catch ArgumentException | Aspose.Cells shape collection get shape by name error handling .NET | What exception is thrown for a non‑existent shape name in Aspose.Cells
// Tags: Aspose.Cells retrieve TextBox shape by name | C# exception handling for missing worksheet shape | Aspose.Cells shape collection ArgumentException handling | update TextBox text in Excel using Aspose.Cells | load and save workbook with error handling Aspose.Cells | validate shape existence before modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, attempts to get a TextBox shape by its assigned name, updates the shape's text, and saves the file. It includes try‑catch blocks to handle file‑load failures, missing shape names (ArgumentException), and any other errors during shape manipulation or saving.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string textboxName = "MyTextBox";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        try
        {
            // Retrieve the TextBox shape by its name
            Shape textbox = sheet.Shapes[textboxName];

            // Example operation on the found TextBox
            textbox.Text = "Updated text";
        }
        catch (ArgumentException ex)
        {
            // Handle the case where the specified TextBox name does not exist
            Console.WriteLine($"Textbox '{textboxName}' not found: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General exception handling for shape operations
            Console.WriteLine($"Error processing textbox: {ex.Message}");
        }

        try
        {
            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
