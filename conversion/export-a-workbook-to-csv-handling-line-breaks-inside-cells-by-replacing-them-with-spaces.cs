// Title: Export Workbook to CSV with Line Breaks Replaced by Spaces using Aspose.Cells (C#)
// Description: Creates a workbook, inserts cells containing LF, CRLF, or CR, replaces newline characters in all string cells with a single space, and saves the result as a CSV file using TxtSaveOptions.
// Keywords: Aspose.Cells | C# | CSV export | remove line breaks | replace newline characters | TxtSaveOptions | SaveFormat.Csv | Excel to CSV conversion | clean cell values | CRLF handling
// Common Searches: Aspose.Cells replace line breaks before CSV export | C# export Excel to CSV without newlines | remove CRLF from cells when saving as CSV | how to clean multi‑line text for CSV using Aspose.Cells | save workbook as CSV with spaces instead of line breaks
// Developer Intent: Strip newline characters from every string cell and generate a CSV file that contains only single‑line values.
// Use Cases: Prepare Excel data for systems that reject line breaks in CSV fields | Create clean, single‑line CSV reports from worksheets with wrapped text | Sanitize user‑generated comments or notes before bulk data export
// AI Prompts: Show how to replace line breaks with a custom delimiter (e.g., '|') instead of a space. | Demonstrate applying the newline removal only to selected columns while exporting to CSV. | Explain how to preserve quoted fields using CsvSaveOptions after cleaning newline characters.

using System;
using Aspose.Cells;

// Creates a workbook, inserts cells containing LF, CRLF, or CR, replaces newline characters in all string cells with a single space, and saves the result as a CSV file using TxtSaveOptions.
class ExportCsvWithLineBreaksHandled
{
    static void Main()
    {
        // Create a new workbook and add sample data containing line breaks
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("First line\nSecond line");
        sheet.Cells["B1"].PutValue("Hello\r\nWorld");
        sheet.Cells["C1"].PutValue("NoBreaksHere");

        // Iterate through all worksheets and cells to replace line breaks with spaces
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;
            foreach (Cell cell in cells)
            {
                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string original = cell.StringValue;
                    if (original.Contains("\n") || original.Contains("\r"))
                    {
                        // Replace CRLF, LF, and CR with a single space
                        string cleaned = original.Replace("\r\n", " ")
                                                .Replace("\n", " ")
                                                .Replace("\r", " ");
                        cell.PutValue(cleaned);
                    }
                }
            }
        }

        // Configure CSV save options (using TxtSaveOptions)
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ','; // default separator for CSV

        // Save the workbook to a CSV file using the Save(string, SaveOptions) method
        workbook.Save("output.csv", csvOptions);
    }
}
