// Title: Export Excel to Pipe‑Delimited TXT with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, configures TxtSaveOptions to use the pipe character (|) as the field separator, and saves the workbook as a pipe‑delimited text file.
// Keywords: Aspose.Cells C# pipe delimited export | TxtSaveOptions custom separator | Excel to pipe separated text .NET | save workbook as txt Aspose | C# export Excel as | delimited file
// Common Searches: Aspose.Cells export Excel as pipe delimited file | C# TxtSaveOptions set separator string | How to save Excel as .txt with custom delimiter using Aspose | Convert .xlsx to pipe separated txt C# | Save workbook as txt with | delimiter Aspose.Cells
// Developer Intent: Export an Excel workbook to a pipe‑separated text file using Aspose.Cells.
// Use Cases: Generate pipe‑delimited reports for data pipelines that require ‘|’ as the field separator. | Create text files compatible with legacy systems that only accept pipe‑separated values. | Export workbook data for bulk import into databases that support pipe‑delimited formats.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to a pipe‑delimited .txt with UTF‑8 encoding and include column headers. | Explain how to batch‑process a folder of Excel files into pipe‑delimited text files with Aspose.Cells. | Show how to configure TxtSaveOptions to skip empty rows when exporting to a pipe‑separated text file.

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures TxtSaveOptions to use the pipe character (|) as the field separator, and saves the workbook as a pipe‑delimited text file.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Set up text save options with a pipe (|) as the delimiter
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.SeparatorString = "|";

        // Save the workbook as a pipe‑delimited TXT file
        string destinationPath = "output.txt";
        workbook.Save(destinationPath, txtOptions);

        Console.WriteLine($"Workbook successfully saved to {destinationPath}");
    }
}
