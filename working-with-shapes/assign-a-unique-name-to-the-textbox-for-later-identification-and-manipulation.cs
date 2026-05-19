using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextboxNamingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet
            // Parameters: upper left row, upper left column, width, height (in pixels)
            int textboxIndex = worksheet.TextBoxes.Add(2, 1, 160, 200);

            // Retrieve the newly added textbox
            TextBox textbox = worksheet.TextBoxes[textboxIndex];

            // Assign a unique name for later identification
            textbox.Name = "UniqueTextBox_001";

            // (Optional) Set some content to verify later
            textbox.Text = "This textbox can be accessed by its unique name.";

            // Demonstrate accessing the textbox later using its name
            TextBox retrievedTextbox = worksheet.TextBoxes["UniqueTextBox_001"];
            if (retrievedTextbox != null)
            {
                Console.WriteLine("Retrieved TextBox Name: " + retrievedTextbox.Name);
                Console.WriteLine("Retrieved TextBox Text: " + retrievedTextbox.Text);
            }

            // Save the workbook
            workbook.Save("TextboxNamingDemo.xlsx");
        }
    }
}