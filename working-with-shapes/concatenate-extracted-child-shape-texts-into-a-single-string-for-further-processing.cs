// Title: Extract and Concatenate Text from All Shapes (including grouped) in an Excel Worksheet – Aspose.Cells for .NET (C#)
// Description: Loads a workbook, walks through every shape on the first worksheet, recursively reads text from group shapes and their children (preferring TextBody.Text, falling back to Shape.Text), builds a single space‑separated string, prints it, and saves the file.
// Keywords: Aspose.Cells | C# | shape text extraction | grouped shapes | concatenate shape texts | Excel worksheet shapes | TextBody | Shape.Text | StringBuilder | Excel automation
// Common Searches: Aspose.Cells get text from grouped shapes | C# concatenate all shape texts in Excel | How to read shape TextBody with Aspose.Cells | Iterate worksheet shapes Aspose.Cells .NET | Extract shape comments from Excel using Aspose
// Developer Intent: Retrieve the textual content of every shape on a worksheet—including nested shapes in groups—and combine it into one string.
// Use Cases: Generate a summary of all annotations, labels, and comments embedded in worksheet shapes for reporting. | Create a searchable index of shape content for document management or compliance audits. | Feed the combined text into downstream processes such as language detection, keyword extraction, or AI summarization.
// AI Prompts: Write a C# method that extracts text from every shape in a worksheet using Aspose.Cells, handling grouped shapes, and returns a concatenated string. | Modify the sample to separate each shape's text with a newline instead of a space. | Explain why TextBody.Text should be prioritized over Shape.Text when reading shape content with Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, walks through every shape on the first worksheet, recursively reads text from group shapes and their children (preferring TextBody.Text, falling back to Shape.Text), builds a single space‑separated string, prints it, and saves the file.
class ConcatenateShapeTexts
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // StringBuilder to accumulate texts from all shapes
        StringBuilder concatenated = new StringBuilder();

        // Iterate through each shape on the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // If the shape is a group, process its child shapes
            if (shape is GroupShape groupShape)
            {
                foreach (Shape childShape in groupShape.GetGroupedShapes())
                {
                    AppendShapeText(childShape, concatenated);
                }
            }
            else
            {
                AppendShapeText(shape, concatenated);
            }
        }

        // Resulting concatenated string
        string result = concatenated.ToString().Trim();
        Console.WriteLine("Concatenated Shape Texts: " + result);

        // Save the workbook if any modifications were made
        workbook.Save("output.xlsx");
    }

    // Helper method to extract text from a shape and append it to the StringBuilder
    private static void AppendShapeText(Shape shape, StringBuilder sb)
    {
        // Prefer TextBody.Text (rich text) if available; otherwise use Shape.Text
        string text = null;

        if (shape.TextBody != null && !string.IsNullOrEmpty(shape.TextBody.Text))
        {
            text = shape.TextBody.Text;
        }
        else if (!string.IsNullOrEmpty(shape.Text))
        {
            text = shape.Text;
        }

        if (!string.IsNullOrEmpty(text))
        {
            sb.Append(text);
            sb.Append(" "); // Add a space as a separator between shape texts
        }
    }
}
