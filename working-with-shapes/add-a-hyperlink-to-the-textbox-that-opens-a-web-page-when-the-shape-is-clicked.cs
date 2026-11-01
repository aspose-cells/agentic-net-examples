// Title: How to add a clickable hyperlink to a textbox shape in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, insert a textbox shape at row 2 column 2, set its display text, assign a hyperlink to https://www.example.com, and save the file as HyperlinkedTextbox.xlsx using Aspose.Cells in C#. | Update an existing worksheet to add a textbox shape that opens a specified URL when clicked, then persist the changes with Aspose.Cells for .NET. | Write C# code that sets the Hyperlink.Address property of a Shape object to link a textbox shape to an external web page in an Excel workbook.
// Common Searches: Aspose.Cells C# add hyperlink to textbox shape in Excel workbook | How to make a textbox shape open a web page when clicked using Aspose.Cells .NET | Set shape click action to URL in Aspose.Cells C# example | Create clickable textbox in Excel with Aspose.Cells library
// Tags: Aspose.Cells add textbox shape hyperlink | C# set shape Hyperlink.Address Aspose.Cells | Excel workbook save with linked shape .NET | Aspose.Cells clickable textbox example | hyperlink assignment to shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Demonstrates creating a workbook, adding a textbox shape to the first worksheet, setting its text, assigning a hyperlink to https://www.example.com via the Hyperlink.Address property, and saving the file as HyperlinkedTextbox.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape (row 2, column 2, offsets 0,0, height 50, width 200)
                Shape textBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 50, 200);

                // Set the displayed text inside the textbox
                textBox.Text = "Click here to open the web page";

                // Assign a hyperlink to the textbox shape
                textBox.Hyperlink.Address = "https://www.example.com";

                // Define output file name
                string outputPath = "HyperlinkedTextbox.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
