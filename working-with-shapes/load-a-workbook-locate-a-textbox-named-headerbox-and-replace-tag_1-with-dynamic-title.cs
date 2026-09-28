// Title: Replace a placeholder tag in a TextBox named HeaderBox in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, finds the Shape whose Name is "HeaderBox", replaces the string "<TAG_1>" in its Text property with a variable title, and saves the workbook to a new location. | Create a reusable C# method that takes a workbook path, a textbox name, a placeholder token, and a replacement string, then updates the textbox text using Aspose.Cells and writes the result to an output file. | Generate a C# example that iterates through all worksheets, locates a TextBox shape, performs multiple placeholder replacements (e.g., <TAG_1>, <TAG_2>) in its content, and saves the modified workbook.
// Common Searches: Aspose.Cells C# replace placeholder text in a specific TextBox shape of an Excel workbook | Find and edit a named TextBox in an .xlsx file using Aspose.Cells for .NET | How to programmatically update <TAG_1> token inside a textbox named HeaderBox with Aspose.Cells | C# iterate worksheets to modify shape text in Excel using Aspose.Cells
// Tags: replace placeholder in Excel textbox Aspose.Cells | search shape by name Aspose.Cells .NET | update TextBox shape text Aspose.Cells | modify shape content in .xlsx using C# | dynamic title insertion into Excel textbox

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an input Excel file, searches all worksheets for a TextBox shape named "HeaderBox", replaces the "<TAG_1>" placeholder with a dynamic title, and saves the updated workbook to the specified output path.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Dynamic title to replace the placeholder tag
        string dynamicTitle = "Quarterly Report";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            bool headerBoxFound = false;

            // Search for the TextBox named "HeaderBox"
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check only the name; assume it's the intended TextBox
                    if (shape.Name == "HeaderBox")
                    {
                        // Replace placeholder with dynamic title
                        string currentText = shape.Text;
                        string updatedText = currentText.Replace("<TAG_1>", dynamicTitle);
                        shape.Text = updatedText;

                        headerBoxFound = true;
                        break; // Exit inner loop
                    }
                }

                if (headerBoxFound)
                    break; // Exit outer loop
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
