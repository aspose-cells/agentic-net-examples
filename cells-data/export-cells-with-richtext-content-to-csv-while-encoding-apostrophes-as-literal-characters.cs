using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExportRichTextCsv
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add a cell with rich‑text content.
                // The leading apostrophe is a text qualifier (QuotePrefix) in Excel.
                cells["A1"].PutValue("'Rich Text Example");

                // Optionally apply rich‑text formatting to parts of the string.
                // Here we make the whole cell bold as a fallback.
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;
                cells["A1"].SetStyle(boldStyle);

                // Configure TXT (CSV) save options.
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ',',                 // Use comma as CSV delimiter
                    Encoding = Encoding.UTF8,        // UTF‑8 encoding
                    ExportQuotePrefix = true,        // Export the leading apostrophe as a literal character
                    QuoteType = TxtValueQuoteType.Minimum // Quote only when necessary
                };

                // Define output file path
                string outputPath = "RichTextExport.csv";

                // Save the workbook as CSV. The apostrophe will appear in the output file.
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}