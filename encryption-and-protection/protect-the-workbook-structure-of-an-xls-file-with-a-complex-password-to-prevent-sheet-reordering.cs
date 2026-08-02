// Title: Protect Excel Workbook Structure with a Complex Password using Aspose.Cells for .NET (C#)
// Description: Load an existing XLS file, apply structure‑only protection with a strong password to stop sheet reordering, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | workbook structure protection | complex password | prevent sheet reordering | XLS to XLSX conversion | Excel security .NET | ProtectionType.Structure
// Common Searches: Aspose.Cells protect workbook structure C# | Set strong password for Excel workbook using .NET | Prevent sheet order changes in XLS file Aspose.Cells | Convert protected XLS to XLSX with Aspose.Cells | Apply only structure protection to Excel via C#
// Developer Intent: Add structure‑only protection with a strong password to an existing XLS workbook and export it as XLSX.
// Use Cases: Distribute a template where users must keep the original sheet order. | Secure legacy XLS reports before sharing them in the more compatible XLSX format. | Lock the layout of a multi‑sheet financial model to enforce a fixed sequence across a team.
// AI Prompts: Generate C# code that protects both workbook structure and windows with a complex password using Aspose.Cells. | Show how to verify if a workbook already has structure protection before applying a new password. | Explain the steps to change the password of an already protected workbook structure with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load an existing XLS file, apply structure‑only protection with a strong password to stop sheet reordering, and save the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "protected_output.xlsx";
        const string complexPassword = "C0mpl3xP@ssw0rd!#2026";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook; Auto format detection handles .xls files
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            using (Workbook workbook = new Workbook(inputPath, loadOptions))
            {
                // Protect only the workbook structure (prevents sheet reordering)
                workbook.Protect(ProtectionType.Structure, complexPassword);

                // Save the protected workbook in XLSX format (widely supported)
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine($"Workbook protected and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
