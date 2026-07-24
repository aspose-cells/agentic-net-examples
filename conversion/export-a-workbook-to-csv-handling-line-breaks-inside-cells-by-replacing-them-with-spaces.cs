// Title: Export Aspose.Cells Workbook to CSV with Line‑Breaks Replaced by Spaces (C#)
// Description: Creates a workbook, inserts sample text containing \n and \r\n, scans every string cell to substitute newline characters with a single space, and saves the result as a CSV file using TxtSaveOptions and a comma delimiter.
// Keywords: Aspose.Cells CSV export | C# remove newline from cells | replace line break Aspose | TxtSaveOptions CSV | Excel to flat file | single‑line CSV output | data cleaning before export
// Common Searches: Aspose.Cells replace newline before CSV | C# export Excel to CSV without line breaks | remove \r\n from cells Aspose.Cells | CSV export with space instead of line break | how to clean multi‑line cells for CSV
// Developer Intent: Strip newline characters from all string cells and generate a CSV file that contains only single‑line values.
// Use Cases: Prepare Excel reports for systems that reject embedded line breaks in CSV files. | Normalize user‑entered multi‑line data before feeding it into a data‑migration pipeline. | Create clean, one‑row‑per‑record CSVs for downstream analytics or import tools.
// AI Prompts: Write C# code with Aspose.Cells that replaces every \r, \n, or \r\n in string cells with a space and then saves the workbook as a CSV. | Show how to configure TxtSaveOptions for CSV export while ensuring no line‑break characters appear in the output. | Explain how to adapt the example to use a semicolon delimiter and keep numeric formatting unchanged during CSV generation.

using System;
using Aspose.Cells;

// Creates a workbook, inserts sample text containing \n and \r\n, scans every string cell to substitute newline characters with a single space, and saves the result as a CSV file using TxtSaveOptions and a comma delimiter.
class ExportCsvWithLineBreakHandling
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data containing line breaks
        sheet.Cells["A1"].PutValue("First line\nSecond line");
        sheet.Cells["B1"].PutValue("Hello\r\nWorld");
        sheet.Cells["C1"].PutValue("NoBreaksHere");

        // Replace line breaks in all string cells with a space
        foreach (Cell cell in sheet.Cells)
        {
            if (cell.Type == CellValueType.IsString)
            {
                string text = cell.StringValue;
                if (text.Contains("\n") || text.Contains("\r"))
                {
                    string cleaned = text.Replace("\r\n", " ")
                                         .Replace("\n", " ")
                                         .Replace("\r", " ");
                    cell.PutValue(cleaned);
                }
            }
        }

        // Set up CSV save options
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ','; // default separator

        // Export the workbook to CSV
        workbook.Save("output.csv", saveOptions);
    }
}
