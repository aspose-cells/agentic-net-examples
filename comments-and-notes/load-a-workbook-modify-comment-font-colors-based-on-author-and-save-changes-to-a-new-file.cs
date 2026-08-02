// Title: Set Excel Comment Font Color by Author with Aspose.Cells for .NET (C#)
// Description: The sample loads an existing Excel file, iterates every worksheet and its comment collection, assigns a specific font color according to the comment’s author (e.g., Alice → Red, Bob → Blue, others → Black), and writes the modified workbook to a new file.
// Keywords: Aspose.Cells | C# Excel automation | comment font color | author based color | modify Excel comments | Workbook.Save | Excel comment example | Aspose.Cells tutorial | C# Excel library | Excel styling by author
// Common Searches: Aspose.Cells change comment color by author C# | C# set Excel comment font color programmatically | Iterate Excel comments with Aspose.Cells | Save workbook after editing comments .NET | How to color-code Excel comments using Aspose
// Developer Intent: Automatically recolor comment text in an Excel workbook according to each author and persist the changes.
// Use Cases: Visually differentiate reviewer remarks in generated reports by applying distinct colors per contributor. | Prepare a shared workbook where comment colors reflect team roles before distribution. | Implement a post‑processing step that highlights critical feedback by coloring comments from designated users.
// AI Prompts: Write C# code using Aspose.Cells that loops through all worksheets, reads each comment’s Author property, sets Font.Color based on a custom mapping, and saves the file. | Show how to extend the author‑to‑color dictionary to include additional users and apply it across the entire workbook. | Explain the steps to access CommentCollection, modify Font.Color, and persist the workbook with Aspose.Cells in .NET. | Provide a GitHub‑style snippet that demonstrates loading a workbook, updating comment colors, and saving to a new path.

using System;
using System.Drawing;
using Aspose.Cells;

namespace CommentFontColorModifier
{
    // The sample loads an existing Excel file, iterates every worksheet and its comment collection, assigns a specific font color according to the comment’s author (e.g., Alice → Red, Bob → Blue, others → Black), and writes the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string sourcePath = "InputWorkbook.xlsx";

            // Load the workbook using the provided constructor (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the comments collection of the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment in the collection
                for (int i = 0; i < comments.Count; i++)
                {
                    Comment comment = comments[i];

                    // Determine the font color based on the comment author
                    // Example mapping: Alice -> Red, Bob -> Blue, others -> Black
                    Color fontColor;
                    switch (comment.Author?.Trim())
                    {
                        case "Alice":
                            fontColor = Color.Red;
                            break;
                        case "Bob":
                            fontColor = Color.Blue;
                            break;
                        default:
                            fontColor = Color.Black;
                            break;
                    }

                    // Apply the selected color to the comment's font
                    comment.Font.Color = fontColor;
                }
            }

            // Save the modified workbook to a new file (save rule)
            string outputPath = "ModifiedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
