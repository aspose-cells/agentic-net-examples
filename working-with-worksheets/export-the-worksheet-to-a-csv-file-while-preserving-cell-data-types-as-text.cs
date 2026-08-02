// Title: Export a Worksheet to CSV as Displayed Text with Aspose.Cells (C#)
// Description: Shows how to save an Aspose.Cells worksheet as a CSV file while writing each cell exactly as it appears on screen. The code applies a date style, configures TxtSaveOptions with a comma delimiter, UTF‑8 encoding, the DisplayStyle format strategy, and optional always‑quote for all fields.
// Keywords: Aspose.Cells CSV export C# | DisplayStyle format strategy | TxtSaveOptions CSV | preserve formatting in CSV | export dates as text Aspose.Cells | UTF-8 CSV Aspose.Cells | quote all fields CSV | save worksheet as CSV .NET | cell value as text CSV
// Common Searches: Aspose.Cells export worksheet to CSV preserving formatting | C# TxtSaveOptions DisplayStyle CSV example | How to keep date format when saving Excel to CSV with Aspose | Export boolean and numeric cells as text using Aspose.Cells | CSV output with all fields quoted Aspose.Cells .NET
// Developer Intent: Generate a CSV file from a worksheet where every cell is written as its displayed text rather than its underlying raw value.
// Use Cases: Create CSV reports that must match the on‑screen number precision and date formatting. | Provide UTF‑8 encoded CSV files with all columns quoted for systems that require quoted values. | Convert mixed‑type Excel sheets to CSV for downstream processing while retaining boolean and date representations as formatted text.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to CSV with all values quoted and using the DisplayStyle format strategy. | Show how to configure TxtSaveOptions so that DateTime cells are saved as formatted text in a CSV file. | Explain how to preserve numeric, boolean, and date cell types as text when saving a worksheet to CSV with Aspose.Cells .NET.

using System.Text;
using Aspose.Cells;

// Shows how to save an Aspose.Cells worksheet as a CSV file while writing each cell exactly as it appears on screen. The code applies a date style, configures TxtSaveOptions with a comma delimiter, UTF‑8 encoding, the DisplayStyle format strategy, and optional always‑quote for all fields.
class ExportWorksheetToCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with different data types
        sheet.Cells["A1"].PutValue("Sample Text");          // string
        sheet.Cells["B1"].PutValue(12345);                  // integer
        sheet.Cells["C1"].PutValue(123.456);                // double
        sheet.Cells["D1"].PutValue(true);                   // boolean
        sheet.Cells["E1"].PutValue(System.DateTime.Now);    // DateTime

        // Apply a date format to the DateTime cell so it is exported as formatted text
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // short date format
        sheet.Cells["E1"].SetStyle(dateStyle);

        // Configure TxtSaveOptions for CSV export
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Separator = ',',                // CSV delimiter
            Encoding = Encoding.UTF8,       // Use UTF‑8 encoding
            FormatStrategy = CellValueFormatStrategy.DisplayStyle, // Export displayed text
            QuoteType = TxtValueQuoteType.Always               // Quote all fields (optional)
        };

        // Export the active worksheet to CSV while preserving data types as text
        workbook.Save("WorksheetExportedAsText.csv", saveOptions);
    }
}
