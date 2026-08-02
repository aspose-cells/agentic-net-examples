// Title: Aspose.Cells C# Example – Update TextBox Text and Export Workbook to XLSX
// Description: Demonstrates how to create a workbook, add a TextBox shape, modify its Text property, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# TextBox | modify textbox text | save workbook as xlsx | Excel shape editing .NET | Aspose.Cells example
// Common Searches: Aspose.Cells change textbox content C# | how to edit a TextBox in an Excel file with Aspose.Cells | save workbook after updating shape Aspose.Cells | C# add TextBox to worksheet and set text | export Aspose.Cells workbook to XLSX
// Developer Intent: Programmatically change the text of a TextBox shape in an Excel workbook and write the result to an XLSX file.
// Use Cases: Generate a report template, insert a TextBox with dynamic values, and deliver the final XLSX to users. | Automate bulk updates of captions or notes stored in TextBox shapes across multiple spreadsheets. | Create a new workbook from scratch, add descriptive TextBox elements, and save it for downstream processing.
// AI Prompts: Provide C# code that locates a TextBox by its name in an Aspose.Cells worksheet, updates its Text, and saves the workbook as XLSX. | Show how to iterate over all TextBox shapes in a worksheet, replace their text using a dictionary of replacements, and export the file. | Write an Aspose.Cells example that adds a TextBox at a specific cell range, sets custom text, and saves the workbook to a given path.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a TextBox shape, modify its Text property, and save the file as an XLSX document using Aspose.Cells for .NET.
    public class ChangeTextBoxTextAndSave
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a new TextBox to the worksheet (row, column, width, height in pixels)
                int textBoxIndex = worksheet.TextBoxes.Add(2, 1, 200, 100);

                // Retrieve the TextBox object from the collection
                TextBox textBox = worksheet.TextBoxes[textBoxIndex];

                // Change the text of the TextBox
                textBox.Text = "This is the updated text.";

                // Save the workbook as XLSX
                workbook.Save("UpdatedTextBox.xlsx");
                Console.WriteLine("Workbook saved successfully as UpdatedTextBox.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChangeTextBoxTextAndSave.Run();
        }
    }
}
