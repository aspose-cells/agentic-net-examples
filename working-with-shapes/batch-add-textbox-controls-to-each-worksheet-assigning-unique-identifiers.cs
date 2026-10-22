// Title: How to batch add uniquely named TextBox shapes to every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells Workbook and inserts a TextBox shape with a sequential Name and Text at a fixed cell location. | Create a reusable method that receives a Workbook object, adds a numbered TextBox to each sheet at row 0, column 0, and saves the workbook to a specified file path.
// Common Searches: aspnet c# add textbox shape to each sheet in existing Excel file using Aspose.Cells | batch insert TextBox controls with unique identifiers across all worksheets Aspose.Cells .NET | how to assign sequential names to Excel TextBox shapes programmatically with Aspose.Cells | loop through worksheets and add TextBox at specific cell coordinates Aspose.Cells C# example
// Tags: add textbox shape Aspose.Cells .NET | batch insert shapes Excel workbook | assign sequential names to textbox Aspose.Cells | iterate worksheets add shapes C# | text box positioning Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel file (or creates a new workbook), loops through every worksheet, and adds a TextBox shape at a predefined location. Each TextBox receives a unique Name and Text based on a global counter, ensuring distinct identifiers across sheets. The modified workbook is then saved to the designated output file.
    class BatchAddTextBoxes
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default sheet
                }

                int globalTextBoxCounter = 1;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Position and size for the TextBox (in points)
                    int upperLeftRow = 0;      // 0‑based row index
                    int upperLeftColumn = 0;   // 0‑based column index
                    int top = 5;               // offset from top of the cell
                    int left = 5;              // offset from left of the cell
                    int width = 150;
                    int height = 50;

                    // Add a TextBox shape; the method returns a TextBox object
                    TextBox textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, width, height);

                    // Assign a unique name and default text
                    textBox.Name = $"TextBox_{globalTextBoxCounter}";
                    textBox.Text = $"Box {globalTextBoxCounter}";

                    globalTextBoxCounter++;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
