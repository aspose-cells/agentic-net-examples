// Title: Export Rich‑Text Cells with Leading Apostrophe to CSV – Aspose.Cells for .NET
// Description: Demonstrates how to save a workbook as a CSV file while preserving a leading apostrophe in a rich‑text cell. The example uses TxtSaveOptions to set a comma delimiter, UTF‑8 encoding, always‑quote fields, and enables ExportQuotePrefix so the apostrophe is written as a literal character.
// Keywords: Aspose.Cells | C# CSV export | ExportQuotePrefix | rich text cell | leading apostrophe | TxtSaveOptions | UTF-8 CSV | comma delimiter | quote all fields | quote prefix export
// Common Searches: Aspose.Cells preserve leading apostrophe CSV | ExportQuotePrefix example C# | Save workbook as CSV with quote prefix | TxtSaveOptions CSV settings Aspose.Cells | How to export rich‑text cell to CSV .NET
// Developer Intent: Generate a CSV file from a worksheet that keeps the leading apostrophe in rich‑text cells as a literal character.
// Use Cases: Export data for systems that require the original apostrophe (e.g., identifiers, codes). | Create UTF‑8 CSV files where every column is quoted to avoid parsing errors. | Maintain quote‑prefix markers when converting Excel sheets with user‑entered apostrophes to CSV.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to CSV, ensuring any leading apostrophe in a cell is saved literally. | Show how to configure TxtSaveOptions for a comma delimiter, UTF‑8 encoding, always‑quote fields, and ExportQuotePrefix enabled. | Explain the role of ExportQuotePrefix when converting rich‑text cells to CSV using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to save a workbook as a CSV file while preserving a leading apostrophe in a rich‑text cell. The example uses TxtSaveOptions to set a comma delimiter, UTF‑8 encoding, always‑quote fields, and enables ExportQuotePrefix so the apostrophe is written as a literal character.
class ExportRichTextToCsv
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Cell with rich‑text (quote prefix) – simulate leading apostrophe by including it in the value
            Cell richCell = cells[0, 0];
            richCell.PutValue("'RichText"); // leading apostrophe stored as part of the cell value

            // Normal cell for comparison
            cells[0, 1].PutValue("Normal");

            // Configure TxtSaveOptions to export the apostrophe as a literal character
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',',                 // CSV delimiter
                Encoding = Encoding.UTF8,        // Desired encoding
                ExportQuotePrefix = true,        // Export the leading apostrophe if present
                QuoteType = TxtValueQuoteType.Always // Quote all fields (optional)
            };

            // Save the workbook as a CSV file using the configured options
            workbook.Save("RichTextExport.csv", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
