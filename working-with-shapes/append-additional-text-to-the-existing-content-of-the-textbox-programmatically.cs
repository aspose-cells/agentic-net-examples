// Title: C# – Append Text to an Aspose.Cells TextBox via FontSettingCollection
// Description: Shows how to create a workbook, insert a TextBox shape, set its initial text, and programmatically append additional characters using FontSettingCollection.AppendText, then save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# TextBox AppendText | FontSettingCollection | Excel shape text manipulation | TextBody property | programmatic textbox update | .NET Excel automation | Add TextBox shape | Aspose.Cells .NET example | Excel textbox editing
// Common Searches: Aspose.Cells AppendText C# example | How to add text to a TextBox shape in Aspose.Cells | FontSettingCollection TextBody usage | Update textbox content Aspose.Cells .NET | C# code to modify Excel textbox
// Developer Intent: Programmatically extend the existing content of a TextBox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Build a greeting textbox by concatenating a static prefix with a dynamic user name. | Append user comments to a pre‑filled notes textbox before distributing a report. | Combine multiple status messages into a single dashboard textbox generated on the fly.
// AI Prompts: Provide C# code that appends several lines of text to an Aspose.Cells TextBox using FontSettingCollection. | Show an example of replacing the entire textbox content with a formatted string and saving the workbook. | Explain how to read the current TextBody text, modify it, and persist the changes in an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, insert a TextBox shape, set its initial text, and programmatically append additional characters using FontSettingCollection.AppendText, then save the workbook as an .xlsx file.
    public class AppendTextToTextBoxDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, height, width (all in pixels)
                TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 100);

                // Access the FontSettingCollection which manages the textbox text
                FontSettingCollection fontSettings = textBox.TextBody;

                // Set the initial text of the textbox
                fontSettings.Text = "Hello, ";

                // Append additional text to the existing content
                fontSettings.AppendText("World!");

                // Optional: display the final text in console
                Console.WriteLine("Final textbox text: " + fontSettings.Text);

                // Save the workbook to a file
                string outputPath = "AppendTextToTextBoxDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AppendTextToTextBoxDemo.Run();
        }
    }
}
