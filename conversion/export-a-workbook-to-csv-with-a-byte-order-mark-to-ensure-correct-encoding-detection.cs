// Title: Export Workbook to CSV with UTF‑8 BOM using AspNet.Cells (C#)
// Description: Creates a workbook, populates sample cells, configures TxtSaveOptions for CSV with Encoding.UTF8 (which automatically writes a UTF‑8 Byte Order Mark), and saves the result as output_with_bom.csv.
// Keywords: Aspose.Cells | C# | CSV export | UTF-8 BOM | TxtSaveOptions | SaveFormat.Csv | Encoding.UTF8 | Byte Order Mark | CSV encoding detection | Excel to CSV
// Common Searches: Aspose.Cells C# export CSV with BOM | How to add UTF-8 BOM to CSV using Aspose.Cells | Save workbook as CSV UTF-8 BOM C# | TxtSaveOptions CSV encoding Aspose | CSV file encoding detection Aspose.Cells | Generate CSV with Byte Order Mark in .NET | Aspose.Cells CSV delimiter and BOM | Export Excel to CSV for international characters
// Developer Intent: Create a CSV file from an Aspose.Cells workbook that includes a UTF‑8 Byte Order Mark so downstream applications reliably detect the encoding.
// Use Cases: Produce CSV reports for systems that require an explicit UTF‑8 BOM for proper character rendering. | Export multilingual data to CSV for opening in Excel or other editors without garbled text. | Integrate CSV generation into a web API where clients expect a BOM‑prefixed UTF‑8 file.
// AI Prompts: Show how to change the code to use UTF-16LE with a BOM for CSV export in Aspose.Cells. | Provide an example that streams the CSV to a MemoryStream while preserving the BOM. | Explain how to set a custom delimiter (e.g., semicolon) and still write the UTF‑8 BOM with Aspose.Cells.

using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvBomDemo
{
    // Creates a workbook, populates sample cells, configures TxtSaveOptions for CSV with Encoding.UTF8 (which automatically writes a UTF‑8 Byte Order Mark), and saves the result as output_with_bom.csv.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure CSV save options with UTF-8 encoding (includes BOM)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.Encoding = Encoding.UTF8; // ensures BOM is written

            // Save the workbook as CSV with the specified options
            workbook.Save("output_with_bom.csv", csvOptions);
        }
    }
}
