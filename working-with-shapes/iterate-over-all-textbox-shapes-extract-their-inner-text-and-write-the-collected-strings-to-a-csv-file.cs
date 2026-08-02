// Title: Export All TextBox Shape Text from an Excel Workbook to CSV with Aspose.Cells for .NET
// Description: Loads an .xlsx file, iterates through every worksheet, reads the Text property of each TextBox shape via the TextBoxCollection, and writes the gathered strings to a CSV file with proper quoting while leaving the original workbook unchanged.
// Keywords: Aspose.Cells C# extract TextBox text | Excel TextBox to CSV | read shape text Aspose.Cells | export textbox contents .NET | TextBoxCollection iteration | CSV export Aspose.Cells | C# Excel shape processing
// Common Searches: How to read TextBox shapes from Excel using Aspose.Cells C# | Export TextBox contents to CSV with Aspose.Cells .NET | Get all TextBox text in a workbook Aspose.Cells | C# extract shape text from .xlsx file | Aspose.Cells TextBoxCollection example
// Developer Intent: Collect the inner text of every TextBox shape in a workbook and save the results to a CSV file.
// Use Cases: Generate a consolidated list of notes entered in TextBox shapes across multiple sheets. | Create a CSV backup of TextBox content before performing bulk workbook modifications. | Prepare TextBox data for import into a database or analytics pipeline.
// AI Prompts: Write C# code that uses Aspose.Cells to read the Text property of each TextBox on all worksheets and export the values to a CSV file with proper escaping. | Suggest performance tips for processing large workbooks that contain thousands of TextBox shapes when exporting their text to CSV. | Show how to modify the sample to include the worksheet name and TextBox index alongside each text entry in the CSV output.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an .xlsx file, iterates through every worksheet, reads the Text property of each TextBox shape via the TextBoxCollection, and writes the gathered strings to a CSV file with proper quoting while leaving the original workbook unchanged.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // List to hold the text from each TextBox
        List<string> textboxTexts = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of TextBox shapes on the current worksheet
            TextBoxCollection textBoxCollection = sheet.TextBoxes;

            // Iterate over each TextBox and collect its inner text
            foreach (TextBox tb in textBoxCollection)
            {
                textboxTexts.Add(tb.Text);
            }
        }

        // Write the collected texts to a CSV file
        string csvFile = "textbox_texts.csv";
        using (StreamWriter writer = new StreamWriter(csvFile))
        {
            // Optional header row
            writer.WriteLine("TextBoxText");

            foreach (string txt in textboxTexts)
            {
                // Escape double quotes for CSV compliance
                string escaped = txt.Replace("\"", "\"\"");
                writer.WriteLine($"\"{escaped}\"");
            }
        }

        // Save the workbook (unchanged) if you need to persist any modifications
        workbook.Save("output.xlsx");
    }
}
