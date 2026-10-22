// Title: Password‑protect a merged Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load two .xlsx files, copy all worksheets from the second workbook into the first, assign a password via Workbook.Settings.Password, and save the encrypted workbook using Aspose.Cells in C#. Apply workbook‑level encryption after merging worksheets by setting Workbook.Settings.Password before calling Workbook.Save in a .NET application. Combine multiple Excel workbooks and protect the resulting file with a custom password using Aspose.Cells' Settings.Password property in C#.
// Common Searches: how to set a password on a merged Excel workbook using Aspose.Cells C# Aspose.Cells encrypt workbook after merging multiple .xlsx files C# code to protect merged workbook with password in Aspose.Cells Workbook.Settings.Password example after adding worksheets Aspose.Cells save merged Excel file as password‑protected using Aspose.Cells for .NET
// Tags: merge worksheets Aspose.Cells C# | Workbook.Settings.Password usage | encrypt Excel workbook .NET | password protect merged workbook Aspose.Cells | combine Excel files with Aspose.Cells and set password

using Aspose.Cells;
using System;
using System.IO;

// The program loads two Excel files, merges all worksheets from the second into the first, sets a password via Workbook.Settings.Password, and saves the result as a password‑protected workbook (merged_protected.xlsx) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to source workbooks
            string sourcePath1 = "source1.xlsx";
            string sourcePath2 = "source2.xlsx";

            // Verify that source files exist
            if (!File.Exists(sourcePath1) || !File.Exists(sourcePath2))
            {
                Console.WriteLine("One or both source files are missing.");
                return;
            }

            // Load the first workbook (source)
            Workbook wb1 = new Workbook(sourcePath1);

            // Load the second workbook (source)
            Workbook wb2 = new Workbook(sourcePath2);

            // Merge all worksheets from wb2 into wb1
            foreach (Worksheet sheet in wb2.Worksheets)
            {
                // Add a copy of each sheet to the first workbook using the sheet name
                wb1.Worksheets.AddCopy(sheet.Name);
            }

            // Encrypt the merged workbook with a password
            wb1.Settings.Password = "Secret123";

            // Save the encrypted workbook
            string outputPath = "merged_protected.xlsx";
            wb1.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
