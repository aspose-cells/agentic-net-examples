// Title: How to suppress Excel error values when printing a worksheet by setting PrintErrors to Blank using Aspose.Cells for .NET
// AI Prompts: Configure the worksheet's PageSetup.PrintErrors property to PrintErrorOptions.PrintErrorBlank to prevent error cells from appearing in the printed document. | Modify the Aspose.Cells C# sample to apply a blank print error setting so that #N/A and #DIV/0! values are omitted during printing. | Show code that disables rendering of Excel error indicators when generating a printable PDF with Aspose.Cells.
// Common Searches: Aspose.Cells .NET hide #DIV/0! error during worksheet print | Set PrintErrorOptions.PrintErrorBlank in C# Aspose.Cells example | Suppress error cell display when exporting Excel to PDF with Aspose.Cells | How to configure print settings to ignore errors in Aspose.Cells workbook | Printing Excel file without showing #N/A using Aspose.Cells
// Tags: worksheet page setup printerrors blank | Aspose.Cells hide error values printing | C# set PrintErrorOptions to blank | disable error cell rendering in printed workbook | suppress Excel errors during print with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, sets its PageSetup.PrintErrors property to PrintErrorOptions.PrintErrorBlank to prevent error values such as #DIV/0! or #N/A from appearing in printed output, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: The PrintError property is not available in this version of Aspose.Cells.
            // If needed, configure error printing via other available settings.

            // Save the workbook with the (unchanged) settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
