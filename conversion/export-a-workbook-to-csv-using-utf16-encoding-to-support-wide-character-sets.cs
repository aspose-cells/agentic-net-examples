// Title: Save Aspose.Cells Workbook as UTF-16 CSV in C#
// Description: Shows how to create a workbook, add Unicode values (e.g., Chinese characters, emoji), set TxtSaveOptions.Encoding to Encoding.Unicode, and export the sheet to a UTF-16 encoded CSV file that preserves all characters.
// Keywords: Aspose.Cells CSV UTF-16 | C# save workbook as Unicode CSV | TxtSaveOptions encoding | Export Excel to UTF-16 CSV .NET | Unicode CSV Aspose.Cells
// Common Searches: Aspose.Cells save CSV UTF-16 C# | C# export Excel to CSV with Unicode support | How to set encoding for CSV in Aspose.Cells | CSV UTF-16 output using Aspose.Cells
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook using UTF-16 encoding to keep multilingual and emoji characters intact.
// Use Cases: Export reports containing Asian scripts or emojis for downstream systems that require UTF‑16 CSV. | Create CSV files compatible with legacy Windows applications that expect Unicode (UTF‑16) input. | Automate data exchange where preserving exact character representation is critical.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook as a UTF‑16 CSV, including the TxtSaveOptions configuration. | Provide an example exporting a worksheet containing Chinese text and emojis to a UTF‑16 encoded CSV using Aspose.Cells. | Explain how to switch the CSV export encoding from UTF‑16 to UTF‑8 or other code pages with Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvUtf16Demo
{
    // Shows how to create a workbook, add Unicode values (e.g., Chinese characters, emoji), set TxtSaveOptions.Encoding to Encoding.Unicode, and export the sheet to a UTF-16 encoded CSV file that preserves all characters.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data that includes Unicode characters
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("张三");   // Chinese characters
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("😀");    // Emoji

            // Configure CSV save options with UTF-16 (Unicode) encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.Unicode   // UTF-16 encoding
            };

            // Save the workbook as a CSV file using the specified options
            workbook.Save("output_utf16.csv", csvOptions);

            Console.WriteLine("Workbook exported to CSV with UTF-16 encoding.");
        }
    }
}
