using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLS workbook that may contain Office Add‑Ins (interactive controls)
        string sourcePath = "input.xls";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Convert the workbook to PDF. ConversionUtility handles loading and saving internally.
        // Interactive controls are preserved when the format supports them.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Workbook has been successfully converted to PDF.");
    }
}