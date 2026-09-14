// Title: Update the text of a specific TextBox shape in an existing XLSX workbook and save the file using Aspose.Cells for .NET
// AI Prompts: Load an XLSX file with Aspose.Cells, locate a TextBox shape by its name, change its Text property to a custom string, and save the workbook to a new file. | Using C#, retrieve a shape named 'MyTextBox' from the first worksheet, set its content to 'New text for the textbox', and export the workbook as XLSX. | Programmatically modify the text inside a named TextBox in an Excel workbook and write the updated workbook to a different path with Aspose.Cells.
// Common Searches: Aspose.Cells C# change text of a specific TextBox shape in an existing Excel file | how to edit named textbox content in XLSX using Aspose.Cells for .NET | save workbook after updating shape text with Aspose.Cells | retrieve shape by name and set Text property Aspose.Cells C# | update textbox in first worksheet and export as new XLSX file
// Tags: set textbox shape text Aspose.Cells | update named shape content XLSX | modify shape text and save workbook Aspose.Cells | load workbook edit textbox .NET | Aspose.Cells shape text manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing XLSX workbook, finds the TextBox shape named 'MyTextBox' on the first worksheet, updates its Text property, and saves the modified workbook as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string textBoxName = "MyTextBox";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook (XLSX)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the TextBox shape by its name
            Shape textbox = null;
            try
            {
                textbox = sheet.Shapes[textBoxName];
            }
            catch (Exception)
            {
                // Shape not found by name; will handle later
            }

            // If the shape exists, attempt to set its text
            if (textbox != null)
            {
                // Directly set the text; Aspose.Cells will handle the shape type internally
                textbox.Text = "New text for the textbox";
                Console.WriteLine($"Text updated for shape '{textBoxName}'.");
            }
            else
            {
                Console.WriteLine($"Shape '{textBoxName}' not found.");
            }

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
