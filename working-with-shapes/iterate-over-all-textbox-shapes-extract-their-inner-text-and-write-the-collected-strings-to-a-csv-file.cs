using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractTextBoxToCsv
{
    static void Main()
    {
        // Load the workbook (change the path to your source file)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Collect all textbox texts from every worksheet
        List<string> textboxTexts = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each TextBox shape in the current worksheet
            foreach (TextBox tb in sheet.TextBoxes)
            {
                // Store the inner text of the textbox
                textboxTexts.Add(tb.Text);
            }
        }

        // Write the collected strings to a CSV file
        string outputPath = "textbox_texts.csv";
        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            // Optional header
            writer.WriteLine("TextBoxText");

            // Write each textbox text as a separate CSV line, escaping quotes
            foreach (string text in textboxTexts)
            {
                string escaped = text?.Replace("\"", "\"\"");
                writer.WriteLine($"\"{escaped}\"");
            }
        }

        // If you need to save any changes to the workbook, uncomment the line below
        // workbook.Save("output.xlsx");
    }
}