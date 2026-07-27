// Title: C# – Set a 4‑point SpaceAfter for every paragraph in an Aspose.Cells text box
// Description: Creates a workbook, adds a text box with multiple paragraphs, and applies a uniform 4‑point SpaceAfter (points) to each paragraph using Aspose.Cells for .NET. The file is saved and optionally re‑loaded to verify the setting.
// Keywords: Aspose.Cells C# paragraph spacing | SpaceAfter property | TextParagraph SpaceAfterSizeType | Excel shape text box line spacing | set paragraph spacing points
// Common Searches: Aspose.Cells set SpaceAfter points | C# paragraph spacing in Excel shape | how to add space after paragraphs Aspose.Cells | text box line spacing Aspose.Cells .NET | adjust paragraph spacing in Excel workbook
// Developer Intent: Apply a 4‑point SpaceAfter value to each paragraph inside a text box shape in an Excel file using Aspose.Cells for .NET.
// Use Cases: Generate reports with multi‑paragraph annotations that need consistent spacing for readability. | Programmatically format text boxes in Excel templates where precise paragraph spacing is required. | Validate that custom paragraph spacing persists after saving and re‑opening the workbook.
// AI Prompts: Write C# code that creates an Excel workbook, inserts a text box, and sets a 4‑point SpaceAfter for all paragraphs using Aspose.Cells. | Show how to read the SpaceAfter value of a paragraph after loading a saved workbook with Aspose.Cells. | Explain how to assign different SpaceAfter values to individual paragraphs within the same text box in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a text box with multiple paragraphs, and applies a uniform 4‑point SpaceAfter (points) to each paragraph using Aspose.Cells for .NET. The file is saved and optionally re‑loaded to verify the setting.
    public class SetSpaceAfterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
                // Set text with multiple paragraphs (separated by newline)
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs within the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Iterate through each paragraph and set spacing after to 4 points
                foreach (TextParagraph paragraph in paragraphs)
                {
                    paragraph.SpaceAfterSizeType = LineSpaceSizeType.Points;
                    paragraph.SpaceAfter = 4.0;
                }

                // Save the workbook to a file
                string filePath = "SetSpaceAfterDemo.xlsx";
                workbook.Save(filePath);

                // Optional: reload to verify the setting (demonstration purpose)
                if (File.Exists(filePath))
                {
                    Workbook loaded = new Workbook(filePath);
                    if (loaded.Worksheets[0].Shapes.Count > 0)
                    {
                        TextParagraph firstPara = loaded.Worksheets[0].Shapes[0].TextBody.TextParagraphs[0];
                        Console.WriteLine("SpaceAfter of first paragraph: " + firstPara.SpaceAfter);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetSpaceAfterDemo.Run();
        }
    }
}
