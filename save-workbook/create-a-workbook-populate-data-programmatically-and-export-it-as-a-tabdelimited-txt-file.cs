// Title: C# – Export a Populated Aspose.Cells Workbook to a Tab‑Delimited TXT (TSV) Using TxtSaveOptions
// Description: Creates a new Workbook, adds product rows, configures TxtSaveOptions with a tab separator and UTF‑8 encoding, exports only the active sheet, and saves it as a .txt file.
// Keywords: Aspose.Cells | C# | TxtSaveOptions | tab delimited export | TSV file | save workbook as txt | UTF-8 encoding | export active worksheet | sample code | GitHub example
// Common Searches: Aspose.Cells save workbook as tab delimited txt C# | TxtSaveOptions tab separator example | export single worksheet to TSV using Aspose.Cells | C# code to create TSV file from Excel workbook | set UTF-8 encoding for txt export Aspose.Cells
// Developer Intent: Generate a UTF‑8 tab‑delimited text file from a programmatically populated workbook, exporting only the active sheet.
// Use Cases: Produce a sales report and deliver it as a TSV file for database import. | Provide inventory data to a legacy system that only accepts tab‑separated text. | Share worksheet contents with non‑Excel users by exporting the active sheet to a .txt file.
// AI Prompts: Write C# code that uses Aspose.Cells to create a workbook, fill it with sample data, and save it as a UTF‑8 tab‑delimited .txt file, exporting only the active sheet. | Explain how TxtSaveOptions can be configured to change the delimiter, set encoding, and select specific worksheets when exporting to text. | Suggest how to modify the example to export each worksheet to its own separate TSV file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTxtExportExample
{
    // Creates a new Workbook, adds product rows, configures TxtSaveOptions with a tab separator and UTF‑8 encoding, exports only the active sheet, and saves it as a .txt file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data programmatically
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["C1"].PutValue("Price");

            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(5);
            worksheet.Cells["C2"].PutValue(999.99);

            worksheet.Cells["A3"].PutValue("Smartphone");
            worksheet.Cells["B3"].PutValue(12);
            worksheet.Cells["C3"].PutValue(599.49);

            worksheet.Cells["A4"].PutValue("Tablet");
            worksheet.Cells["B4"].PutValue(8);
            worksheet.Cells["C4"].PutValue(399.00);

            // Configure text save options for tab‑delimited output (rule usage)
            TxtSaveOptions txtOptions = new TxtSaveOptions();
            txtOptions.Separator = '\t';               // Tab character as delimiter
            txtOptions.Encoding = System.Text.Encoding.UTF8;
            txtOptions.ExportAllSheets = false;        // Export only the active sheet

            // Save the workbook as a tab‑delimited TXT file (lifecycle rule)
            string outputPath = "ExportedData.txt";
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook exported successfully to '{outputPath}' as a tab‑delimited text file.");
        }
    }
}
