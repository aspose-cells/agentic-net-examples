using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextboxByNameDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet
            int tbIndex = worksheet.TextBoxes.Add(2, 2, 200, 50);
            TextBox tb = worksheet.TextBoxes[tbIndex];

            // Assign a unique name to the textbox
            tb.Name = "MyTextBox";

            // Set initial properties
            tb.Text = "Original Text";
            tb.Font.Size = 12;
            tb.Font.IsBold = false;

            // Retrieve the textbox by its assigned name using the name indexer (rule)
            TextBox retrievedTb = worksheet.TextBoxes["MyTextBox"];
            if (retrievedTb != null)
            {
                // Modify properties programmatically
                retrievedTb.Text = "Updated Text via Name Indexer";
                retrievedTb.Font.IsBold = true;
                retrievedTb.Font.Color = System.Drawing.Color.Blue;
                retrievedTb.Height = 80; // change height
                retrievedTb.Width = 250; // change width
            }

            // Save the workbook (save rule)
            workbook.Save("TextboxByNameDemo.xlsx");
        }
    }
}