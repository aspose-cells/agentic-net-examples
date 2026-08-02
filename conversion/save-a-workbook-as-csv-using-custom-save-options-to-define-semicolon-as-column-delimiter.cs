// Title: Save a Workbook as CSV with a Semicolon Delimiter using Aspose.Cells TxtSaveOptions (C#)
// Description: Creates a workbook, fills it with sample data, configures TxtSaveOptions to use ';' as the column separator and UTF‑8 encoding, then saves the file as output.csv.
// Keywords: Aspose.Cells CSV semicolon delimiter | TxtSaveOptions separator C# | custom CSV delimiter Aspose.Cells | export workbook to CSV UTF-8 | C# Aspose.Cells save as CSV
// Common Searches: Aspose.Cells set semicolon as CSV delimiter C# | TxtSaveOptions separator property example | How to export Excel to CSV with custom separator using Aspose.Cells | C# save workbook as CSV with UTF-8 encoding Aspose
// Developer Intent: Export a workbook to a CSV file where columns are separated by a semicolon instead of a comma.
// Use Cases: Produce CSV files for European locales that default to semicolon as the list separator. | Generate data files for legacy systems that require semicolon‑delimited input. | Create UTF‑8 encoded reports where commas appear in data values, avoiding column misalignment.
// AI Prompts: Give C# code that saves an Aspose.Cells workbook to CSV using a pipe (|) as the column delimiter. | Explain how to set TxtSaveOptions to include a header row and use UTF‑16 encoding when exporting to CSV. | Show how to dynamically choose the CSV delimiter based on user locale with Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Creates a workbook, fills it with sample data, configures TxtSaveOptions to use ';' as the column separator and UTF‑8 encoding, then saves the file as output.csv.
class SaveCsvWithSemicolon
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Configure text save options to use semicolon as the column delimiter
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.Separator = ';';
        saveOptions.Encoding = Encoding.UTF8;

        // Save the workbook as CSV with the custom separator
        workbook.Save("output.csv", saveOptions);
    }
}
