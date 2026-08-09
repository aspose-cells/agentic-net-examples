// Title: C# – Save Aspose.Cells Workbook as CSV with Semicolon Delimiter
// Description: Creates a workbook, fills cells A1‑B3, configures TxtSaveOptions to use a semicolon (;) as the column separator and UTF‑8 encoding, then saves the file as "output_semicolon.csv".
// Keywords: Aspose.Cells | C# | .NET | CSV export | semicolon delimiter | TxtSaveOptions | custom column separator | UTF-8 encoding | European locale CSV | Excel to CSV conversion
// Common Searches: Aspose.Cells save CSV with semicolon | C# TxtSaveOptions separator property | export Excel to CSV using custom delimiter .NET | how to set CSV delimiter in Aspose.Cells | UTF-8 CSV export Aspose.Cells C#
// Developer Intent: Export an Aspose.Cells workbook to a CSV file using a semicolon as the column separator.
// Use Cases: Generate CSV reports for European markets where ';' is the default list separator. | Produce data files compatible with legacy systems that require semicolon‑delimited CSV. | Export multilingual data with UTF‑8 encoding while controlling the delimiter.
// AI Prompts: Write C# code to save an Aspose.Cells workbook as a pipe‑delimited CSV using TxtSaveOptions. | Explain how to import a semicolon‑delimited CSV back into an Aspose.Cells workbook in C#. | Show how to configure ISO‑8859‑1 encoding with a custom delimiter when saving a workbook as CSV.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvSemicolonDemo
{
    // Creates a workbook, fills cells A1‑B3, configures TxtSaveOptions to use a semicolon (;) as the column separator and UTF‑8 encoding, then saves the file as "output_semicolon.csv".
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Configure text save options for CSV with semicolon delimiter
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ';',               // Use semicolon as column separator
                Encoding = Encoding.UTF8      // Optional: set encoding
            };

            // Save the workbook as a CSV file using the custom options
            string outputPath = "output_semicolon.csv";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with semicolon delimiter.");
        }
    }
}
