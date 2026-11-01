// Title: How to programmatically append custom text to an existing TextBox shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that locates the first TextBox shape on a worksheet and appends a specified string to its current text using Aspose.Cells. | Create a version of the sample that iterates through all TextBox shapes in a worksheet and appends the same text to each one. | Write code that reads the text to be appended from a command‑line argument or external file and updates the TextBox content accordingly.
// Common Searches: aspnet cells c# append text to textbox shape in existing excel file | how to modify textbox content in excel using Aspose.Cells library | c# iterate worksheet shapes and update textbox text with Aspose.Cells | append string to Excel textbox without overwriting existing text Aspose.Cells | read append text from file and update Excel textbox Aspose.Cells C#
// Tags: append text to Excel TextBox Aspose.Cells | modify TextBox shape content .NET | iterate worksheet shapes Aspose.Cells | save workbook after textbox update Aspose.Cells | read append string from external file C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads Input.xlsx, finds the first TextBox shape on the first worksheet, appends " - Appended text." to its existing Text, and saves the modified workbook as Output.xlsx.
class AppendTextToTextbox
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (modify as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet to find a TextBox
            foreach (Shape shape in sheet.Shapes)
            {
                // Use pattern matching to identify a TextBox shape
                if (shape is TextBox textbox)
                {
                    // Append additional text to the existing content
                    string additionalText = " - Appended text.";
                    textbox.Text = textbox.Text + additionalText;

                    // If only the first textbox needs modification, exit the loop
                    break;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
