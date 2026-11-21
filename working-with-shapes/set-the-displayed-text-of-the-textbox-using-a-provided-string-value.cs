// Title: Setting the displayed text of a TextBox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: aspnetc# how to set textbox shape text using Aspose.Cells | add a textbox to an Excel worksheet and assign a string with Aspose.Cells .NET | change the content of an existing TextBox shape in a workbook via Aspose.Cells C# | programmatically update TextBox Text property in Excel using Aspose.Cells for .NET | load workbook, insert textbox, set Text property Aspose.Cells example C#
// Tags: Aspose.Cells textbox shape creation C# | assign Text property to TextBox Aspose.Cells | Excel worksheet shape manipulation .NET | initialize workbook with Aspose.Cells | save workbook after shape changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, adds a TextBox shape to the first worksheet, sets its Text property to a provided string, and saves the workbook to the specified output file.
    class Program
    {
        static void Main()
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Text to display inside the textbox
            string textboxContent = "Your provided string value";

            try
            {
                // Load existing workbook if it exists; otherwise create a new one
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                TextBox textbox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 50);

                // Set the displayed text of the textbox
                textbox.Text = textboxContent;

                // Save the modified workbook
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
