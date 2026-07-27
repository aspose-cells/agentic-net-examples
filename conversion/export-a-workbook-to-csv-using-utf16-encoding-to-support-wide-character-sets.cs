// Title: Export Workbook to UTF‑16 CSV with Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, add multilingual data, configure TxtSaveOptions for UTF‑16 encoding, and save the file as a CSV that preserves Unicode characters such as Chinese and Cyrillic.
// Keywords: Aspose.Cells CSV UTF-16 | C# export workbook to CSV Unicode | TxtSaveOptions encoding Unicode | save Excel as UTF-16 CSV | multilingual CSV export .NET
// Common Searches: Aspose.Cells save CSV with UTF-16 encoding C# | how to export Unicode characters to CSV using Aspose.Cells | TxtSaveOptions CSV Unicode example | C# export Excel to UTF-16 CSV | Aspose.Cells multilingual CSV export
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook using UTF‑16 to retain all Unicode characters.
// Use Cases: Export multilingual reports (e.g., Chinese, Cyrillic) for systems that require UTF‑16 CSV files. | Archive every worksheet of an Excel workbook into a single Unicode‑compatible CSV. | Provide CSV output for legacy applications that only accept UTF‑16 encoded data.
// AI Prompts: Create C# code that uses Aspose.Cells to save a workbook as a UTF‑16 CSV, including options for multiple sheets. | Show how to modify TxtSaveOptions to set a custom delimiter while keeping UTF‑16 encoding. | Explain the steps to read a UTF‑16 CSV produced by Aspose.Cells back into a .NET DataTable.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvUtf16Example
{
    // Demonstrates how to create a workbook, add multilingual data, configure TxtSaveOptions for UTF‑16 encoding, and save the file as a CSV that preserves Unicode characters such as Chinese and Cyrillic.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("张三"); // Unicode characters to demonstrate UTF‑16
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Иван"); // Cyrillic characters

            // Configure CSV save options with UTF‑16 encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.Unicode, // UTF‑16
                ExportAllSheets = true       // Export all worksheets (optional)
            };

            // Save the workbook as CSV using the configured options
            workbook.Save("output_utf16.csv", csvOptions);

            Console.WriteLine("Workbook exported to CSV with UTF‑16 encoding.");
        }
    }
}
