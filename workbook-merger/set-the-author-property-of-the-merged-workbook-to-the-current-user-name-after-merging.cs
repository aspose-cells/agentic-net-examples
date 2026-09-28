// Title: How to set the Author property of a merged Excel workbook to the current Windows user using Aspose.Cells for .NET
// AI Prompts: Merge multiple .xlsx files with Aspose.Cells and assign Environment.UserName to the merged workbook's Author metadata. | Load several workbooks, copy their worksheets into a new workbook, then set mergedWorkbook.Settings.Author to the logged‑in user. | Programmatically update the document author of an Aspose.Cells workbook after merging worksheets in C#.
// Common Searches: Aspose.Cells set workbook author to current user after merging worksheets C# | C# merge multiple Excel files and update document properties with Aspose.Cells | How to assign Environment.UserName to Excel file metadata using Aspose.Cells .NET
// Tags: aspocells merge workbooks set author | c# environment.username excel metadata | aspocells document properties author after merge | excel workbook author property aspocells | merge worksheets update author c#

using System;
using System.IO;
using Aspose.Cells;

// Merges worksheets from multiple source Excel files into a new workbook using Aspose.Cells, sets the merged workbook's Author property to the current Windows user via Environment.UserName, and saves the result as MergedOutput.xlsx.
class MergeWorkbooks
{
    static void Main()
    {
        try
        {
            // Workbook that will hold the merged content
            Workbook mergedWorkbook = new Workbook();

            // Source files to merge
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

            foreach (string file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found: {file}");
                    continue; // Skip missing files
                }

                // Load source workbook
                Workbook src = new Workbook(file);

                // Copy each worksheet by name into the merged workbook
                foreach (Worksheet ws in src.Worksheets)
                {
                    mergedWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Set author property
            mergedWorkbook.Settings.Author = Environment.UserName;

            // Save the merged workbook
            string outputPath = "MergedOutput.xlsx";
            mergedWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
