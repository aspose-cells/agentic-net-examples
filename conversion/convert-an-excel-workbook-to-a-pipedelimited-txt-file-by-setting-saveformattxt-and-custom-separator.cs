// Title: C# – Convert Excel Workbook to Pipe‑Delimited TXT with Aspose.Cells
// Description: Shows how to create or load a workbook, set TxtSaveOptions.SeparatorString to "|" and Encoding to UTF‑8, and save the file as a pipe‑delimited text document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# TxtSaveOptions | pipe delimited txt | custom delimiter | Excel to TXT conversion | SaveFormat.Txt | UTF-8 encoding | export Excel as pipe separated | Aspose.Cells .NET
// Common Searches: Aspose.Cells save workbook as pipe delimited txt | C# set custom separator for TxtSaveOptions | How to export Excel to pipe separated file using Aspose | TxtSaveOptions separator string example | Convert .xlsx to .txt with | delimiter in .NET
// Developer Intent: Create a pipe‑separated text file from an Excel workbook using Aspose.Cells in C#.
// Use Cases: Feed data to legacy applications that require pipe‑separated values. | Generate UTF‑8 delimited files for data pipelines or third‑party services. | Produce custom‑delimited reports directly from Excel without manual CSV conversion. | Prepare input files for mainframe or ETL tools that expect a ‘|’ delimiter.
// AI Prompts: Show me how to change the delimiter to a semicolon in the TxtSaveOptions code. | Provide an example that loads an existing .xlsx file and saves it as a pipe‑delimited TXT while preserving formulas. | Explain how to configure TxtSaveOptions to export only selected columns with a custom delimiter. | Give code to include only column headers in the output pipe‑delimited file.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to create or load a workbook, set TxtSaveOptions.SeparatorString to "|" and Encoding to UTF‑8, and save the file as a pipe‑delimited text document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Configure text save options with a pipe (|) as the delimiter
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.SeparatorString = "|";   // Set custom separator
        txtOptions.Encoding = Encoding.UTF8; // Optional: set encoding

        // Save the workbook as a pipe‑delimited TXT file
        workbook.Save("output.txt", txtOptions);

        Console.WriteLine("Workbook successfully saved as pipe‑delimited TXT.");
    }
}
