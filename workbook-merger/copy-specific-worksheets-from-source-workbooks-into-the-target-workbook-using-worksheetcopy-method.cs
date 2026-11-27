// Title: Copy selected worksheets from multiple Excel workbooks into a single workbook using Aspose.Cells Worksheet.Copy in C#
// Common Searches: c# copy worksheets with prefix 'CopyMe' from several Excel files into one workbook using Aspose.Cells | aspnet merge worksheets from multiple workbooks handling duplicate sheet names with Worksheet.Copy | how to create an empty target workbook and copy sheets from source workbooks in Aspose.Cells .NET | command line tool to copy specific sheets from multiple Excel files into a new file using Aspose.Cells | Aspose.Cells copy first sheet when no matching prefix is found
// Tags: Worksheet.Copy copy worksheets Aspose.Cells | merge multiple workbooks C# Aspose.Cells | select worksheets by name prefix Aspose.Cells | resolve duplicate sheet names Aspose.Cells | create empty target workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetCopyDemo
{
    // The console application parses command‑line arguments, treating all but the last path as source Excel files and the final path as the target workbook. It creates an empty target workbook, removes the default sheet, and iterates through each source workbook. For each source, it selects worksheets whose names start with "CopyMe" (or falls back to the first sheet). Each selected sheet is added to the target workbook via Worksheet.Copy, and name conflicts are resolved by appending a numeric suffix. Finally, the merged workbook is saved to the specified target path.
    class Program
    {
        /// <param name="args">
        /// Command‑line arguments.
        /// All arguments except the last one are treated as source workbook file paths.
        /// The last argument is the target workbook file path where the copied sheets will be saved.
        /// Example:
        ///   WorksheetCopyDemo.exe source1.xlsx source2.xlsx target.xlsx
        /// </param>
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args == null || args.Length < 2)
                {
                    Console.WriteLine("Usage: WorksheetCopyDemo.exe <source1.xlsx> [<source2.xlsx> ...] <target.xlsx>");
                    return;
                }

                // Separate source files from target file
                string targetPath = args[args.Length - 1];
                string[] sourcePaths = new string[args.Length - 1];
                Array.Copy(args, sourcePaths, args.Length - 1);

                // Verify that each source file exists
                foreach (string src in sourcePaths)
                {
                    if (!File.Exists(src))
                    {
                        Console.WriteLine($"Error: Source file not found – \"{src}\"");
                        return;
                    }
                }

                // Create an empty target workbook
                Workbook targetWorkbook = new Workbook();
                // Remove the default empty sheet that Aspose.Cells creates
                if (targetWorkbook.Worksheets.Count > 0)
                {
                    targetWorkbook.Worksheets.Clear();
                }

                // Keep track of sheet names to avoid duplicates in the target workbook
                HashSet<string> existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Process each source workbook
                foreach (string srcPath in sourcePaths)
                {
                    Console.WriteLine($"Processing source workbook: {srcPath}");
                    Workbook sourceWorkbook = new Workbook(srcPath);

                    // Define which worksheets to copy.
                    // For demonstration we copy all worksheets whose name starts with "CopyMe"
                    // Adjust this predicate as needed (e.g., by index, exact name, etc.).
                    List<Worksheet> worksheetsToCopy = new List<Worksheet>();
                    foreach (Worksheet ws in sourceWorkbook.Worksheets)
                    {
                        if (ws.Name.StartsWith("CopyMe", StringComparison.OrdinalIgnoreCase))
                        {
                            worksheetsToCopy.Add(ws);
                        }
                    }

                    // If no worksheet matches the predicate, fall back to copying the first sheet.
                    if (worksheetsToCopy.Count == 0 && sourceWorkbook.Worksheets.Count > 0)
                    {
                        worksheetsToCopy.Add(sourceWorkbook.Worksheets[0]);
                    }

                    // Copy each selected worksheet into the target workbook
                    foreach (Worksheet sourceWs in worksheetsToCopy)
                    {
                        // Add a new blank worksheet to the target workbook
                        int newIndex = targetWorkbook.Worksheets.Add();
                        Worksheet targetWs = targetWorkbook.Worksheets[newIndex];

                        // Use Worksheet.Copy to copy the entire content, formatting, charts, etc.
                        targetWs.Copy(sourceWs);

                        // Resolve name conflicts
                        string desiredName = sourceWs.Name;
                        string finalName = desiredName;
                        int suffix = 1;
                        while (existingNames.Contains(finalName))
                        {
                            finalName = $"{desiredName}_{suffix}";
                            suffix++;
                        }

                        targetWs.Name = finalName;
                        existingNames.Add(finalName);

                        Console.WriteLine($"  Copied worksheet \"{sourceWs.Name}\" as \"{finalName}\"");
                    }
                }

                // Ensure the directory for the target file exists
                string targetDir = Path.GetDirectoryName(Path.GetFullPath(targetPath));
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                // Save the target workbook
                targetWorkbook.Save(targetPath);
                Console.WriteLine($"All selected worksheets have been copied successfully to \"{targetPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
