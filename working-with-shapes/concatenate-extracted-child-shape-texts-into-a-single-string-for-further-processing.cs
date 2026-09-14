// Title: Extract and concatenate text from every shape, including grouped shapes, in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, walks through all worksheets, extracts the Text property from each shape (including TextBox and grouped shapes), and returns a single concatenated string. | Create a reusable method that accepts a Shape object and a StringBuilder, appends any available text from TextBox or generic shapes, and correctly processes shapes inside a group. | Generate a console application example that demonstrates how to collect shape texts from multiple worksheets, combine them using StringBuilder, and output the result.
// Common Searches: how to get text from all shapes in an Excel file using Aspose.Cells C# | concatenate grouped shape text in Aspose.Cells .NET workbook | C# iterate over worksheet shapes and read their Text property Aspose.Cells | extract textbox and autoshape contents from Excel with Aspose.Cells | combine shape texts from multiple sheets into one string Aspose.Cells
// Tags: shape text extraction Aspose.Cells | concatenate Excel shape contents .NET | grouped shape text retrieval C# | iterate worksheet shapes Aspose.Cells | StringBuilder aggregation of shape text

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Text;

// // Loads an Excel workbook, iterates through each worksheet and its shapes, extracts text from TextBox and other shapes (including those inside groups), appends the texts to a StringBuilder, and prints the concatenated result.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var sb = new StringBuilder();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Append text from the shape (including group shapes if they contain text)
                    AppendShapeText(shape, sb);
                }
            }

            // Concatenated result
            string concatenatedText = sb.ToString();
            Console.WriteLine(concatenatedText);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Extracts text from supported shape types and appends it to the StringBuilder
    static void AppendShapeText(Shape shape, StringBuilder sb)
    {
        // TextBox shapes contain a Text property
        if (shape is TextBox textbox)
        {
            sb.Append(textbox.Text);
        }
        // Other shapes (e.g., AutoShape) expose text via the generic Text property
        else if (!string.IsNullOrEmpty(shape.Text))
        {
            sb.Append(shape.Text);
        }
    }
}
