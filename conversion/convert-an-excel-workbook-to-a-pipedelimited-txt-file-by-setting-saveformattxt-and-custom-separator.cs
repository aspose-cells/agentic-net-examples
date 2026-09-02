// Title: How to convert an .xlsx workbook to a pipe‑delimited .txt file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, sets TxtSaveOptions.SeparatorString to "|", and saves the workbook as a pipe‑separated text file. | Show how to set TxtSaveOptions.Encoding to Encoding.UTF8 and define a custom field delimiter for exporting Excel data to a .txt file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# export Excel data as pipe delimited text | Set delimiter for TxtSaveOptions when saving workbook as .txt | Convert .xlsx to pipe delimited .txt using Aspose.Cells .NET | How to specify UTF‑8 encoding for text export with Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions pipe delimiter | C# generate pipe delimited text from Excel | Aspose.Cells text export encoding option | Save workbook as .txt with custom separator | Excel to pipe delimited file using .NET

using System;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, configures TxtSaveOptions to use a pipe (|) as the field separator and UTF‑8 encoding, and saves the workbook as a pipe‑delimited text file.
class ConvertExcelToPipeDelimited
{
    static void Main()
    {
        // Load the source Excel workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Set up text save options with a pipe (|) as the delimiter
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.SeparatorString = "|";
        saveOptions.Encoding = Encoding.UTF8; // optional, ensures UTF‑8 output

        // Save the workbook as a pipe‑delimited TXT file
        string destinationPath = "output.txt";
        workbook.Save(destinationPath, saveOptions);

        Console.WriteLine($"Workbook successfully converted to pipe‑delimited text file: {destinationPath}");
    }
}
