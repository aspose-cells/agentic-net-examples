// Title: Set the Author metadata of a merged Excel workbook to the logged‑in Windows user with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook author after merging multiple Excel files | How to assign the current Windows user as the author property in a merged workbook using Aspose.Cells | Merge several .xlsx files and set author metadata programmatically with Aspose.Cells | Aspose.Cells example for Workbook.Settings.WriteProtection.Author in C# | Use Environment.UserName to set Excel file author with Aspose.Cells
// Tags: Aspose.Cells merge workbooks set author | Workbook.Settings.WriteProtection.Author C# | set Excel file author programmatically Aspose.Cells | merge multiple .xlsx files Aspose.Cells | Environment.UserName Excel author metadata

using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbooksSetAuthor
{
    // The sample loads the first workbook as a base, copies worksheets from additional .xlsx files, sets the merged workbook's author to the current Windows user via Workbook.Settings.WriteProtection.Author, creates the output folder if needed, and saves the combined workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments: input1.xlsx input2.xlsx ... output.xlsx
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: MergeWorkbooksSetAuthor.exe <input1.xlsx> [<input2.xlsx> ...] <output.xlsx>");
                return;
            }

            // The last argument is the output file, the rest are input files
            string outputPath = args[args.Length - 1];
            string[] inputPaths = new string[args.Length - 1];
            Array.Copy(args, inputPaths, args.Length - 1);

            // Validate input files exist
            foreach (var path in inputPaths)
            {
                if (!File.Exists(path))
                {
                    Console.WriteLine($"Error: Input file not found - {path}");
                    return;
                }
            }

            try
            {
                // Load the first workbook as the base merged workbook
                Workbook mergedWorkbook = new Workbook(inputPaths[0]);

                // Merge remaining workbooks
                for (int i = 1; i < inputPaths.Length; i++)
                {
                    Workbook srcWorkbook;
                    try
                    {
                        srcWorkbook = new Workbook(inputPaths[i]);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load workbook '{inputPaths[i]}': {ex.Message}");
                        continue; // skip this workbook and continue with others
                    }

                    foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                    {
                        // Add a copy of each worksheet to the merged workbook by name (compatible overload)
                        mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);
                    }
                }

                // Set the author of the merged workbook to the current user name
                string currentUser = Environment.UserName;
                mergedWorkbook.Settings.WriteProtection.Author = currentUser;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the merged workbook
                mergedWorkbook.Save(outputPath);
                Console.WriteLine($"Merged workbook saved to: {outputPath}");
                Console.WriteLine($"Author property set to: {currentUser}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during merging:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
