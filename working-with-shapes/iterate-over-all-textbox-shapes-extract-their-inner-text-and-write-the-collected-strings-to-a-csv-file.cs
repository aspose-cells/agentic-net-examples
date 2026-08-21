// Title: Extract TextBox Shape Text from an Excel Workbook and Export to CSV with Aspose.Cells for .NET
// Description: Load an Excel file, loop through every worksheet, gather the Text property of each non‑empty TextBox shape, and write the collected strings to a CSV file with proper quoting. The workbook can be saved afterward if needed.
// Keywords: Aspose.Cells TextBox extraction | C# export TextBox to CSV | read shape text Aspose.Cells | .NET Excel TextBox to CSV | iterate worksheets Aspose.Cells | extract textbox contents | Excel shape text export | Aspose.Cells CSV output | workbook TextBox collection | Aspose.Cells shape handling
// Common Searches: how to get text from all TextBox shapes using Aspose.Cells | export Excel TextBox contents to CSV in C# | Aspose.Cells iterate over worksheet TextBoxes | C# read TextBox shape text from workbook | save TextBox values to CSV with Aspose
// Developer Intent: Collect every TextBox shape's inner text from a workbook and write the values to a CSV file.
// Use Cases: Consolidate user comments stored in TextBox shapes across multiple sheets for reporting. | Migrate legacy data entered in TextBoxes to a CSV format for database import. | Create an audit log of TextBox contents before performing bulk edits on the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate all worksheets, extract each TextBox's Text, and export the results to a CSV file with proper escaping. | Show how to filter extracted TextBox text by a keyword before writing to CSV using Aspose.Cells. | Explain performance‑friendly techniques for extracting TextBox contents from large Excel files with Aspose.Cells while keeping memory usage low.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel file, loop through every worksheet, gather the Text property of each non‑empty TextBox shape, and write the collected strings to a CSV file with proper quoting. The workbook can be saved afterward if needed.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Collect the inner text of all TextBox shapes across all worksheets
        List<string> textboxTexts = new List<string>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The TextBoxes property gives access to the TextBox collection of the worksheet
            foreach (TextBox tb in sheet.TextBoxes)
            {
                // Guard against null or empty text
                if (!string.IsNullOrEmpty(tb.Text))
                {
                    textboxTexts.Add(tb.Text);
                }
            }
        }

        // Write the collected strings to a CSV file
        string csvPath = "textbox_texts.csv";
        using (StreamWriter writer = new StreamWriter(csvPath))
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

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
