// Title: Hide columns B through E in a worksheet and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing .xlsx file, hides columns B‑E in the worksheet named 'Sheet1', and writes the result to a new file. | Show how to use Aspose.Cells in C# to conceal a range of columns in a specific worksheet and preserve the hidden state when saving the workbook.
// Common Searches: Aspose.Cells C# hide columns B to E in a specific worksheet | How to hide a range of columns in Excel using Aspose.Cells for .NET | Preserve hidden columns when saving an Excel workbook with Aspose.Cells | C# code to hide multiple columns in an Aspose.Cells worksheet | Aspose.Cells hide column range and save workbook example
// Tags: Aspose.Cells hide column range C# | Aspose.Cells column visibility preservation | Aspose.Cells hide columns in worksheet | C# hide Excel columns Aspose.Cells | Aspose.Cells save workbook after column hide | Aspose.Cells worksheet column manipulation

using Aspose.Cells;
using System;
using System.IO;

// Loads 'input.xlsx', hides columns B‑E (indices 1‑4) in the worksheet 'Sheet1' using Worksheet.Cells.HideColumns, and saves the modified workbook as 'output.xlsx', with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the specific worksheet (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0];

            // Hide columns B through E (indices 1 to 4)
            worksheet.Cells.HideColumns(1, 4);

            // Save the workbook; column visibility is preserved automatically
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
