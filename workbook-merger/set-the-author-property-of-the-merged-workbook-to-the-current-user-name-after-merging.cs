// Title: Merge several Excel workbooks into one file and set the merged workbook’s Author property to the current Windows user using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a list of .xlsx files, copies each worksheet into a new Aspose.Cells workbook, assigns Environment.UserName to BuiltInDocumentProperties.Author, and saves the result. | Write a .NET program that merges multiple Excel workbooks, removes the default sheet, copies worksheets with AddCopy, sets the Author document property to the logged‑in user, and outputs a combined workbook.
// Common Searches: c# aspnet merge multiple excel files and set author property programmatically | how to assign current Windows user as author in merged workbook using Aspose.Cells | copy worksheets from several workbooks into one and update built‑in document properties in .NET | Aspose.Cells set BuiltInDocumentProperties.Author after merging workbooks
// Tags: Aspose.Cells merge workbooks C# | set built‑in document property author Aspose.Cells | copy worksheets using AddCopy Aspose.Cells | Environment.UserName workbook author property | save merged workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // This C# example uses Aspose.Cells to merge several .xlsx workbooks by copying each worksheet into a new workbook, clears the default sheet, sets the merged workbook’s Author built‑in document property to the current Windows user (Environment.UserName), and saves the combined file as an Xlsx document.
class WorkbookMerger
{
    static void Main()
    {
        try
        {
            // Paths of workbooks to merge
            string[] sourceFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new workbook that will hold the merged content
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty worksheet created by the constructor
            mergedWorkbook.Worksheets.Clear();

            // Iterate through each source workbook and copy its worksheets into the merged workbook
            foreach (string filePath in sourceFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue; // Skip missing files
                }

                // Load the source workbook
                Workbook source = new Workbook(filePath);

                // Copy each worksheet from the source workbook
                foreach (Worksheet sheet in source.Worksheets)
                {
                    // Clone the worksheet into the merged workbook using its name
                    mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                }
            }

            // Set the Author property of the merged workbook to the current user name
            mergedWorkbook.BuiltInDocumentProperties.Author = Environment.UserName;

            // Ensure output directory exists
            string outputPath = @"C:\Data\MergedWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the merged workbook
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged and saved to '{outputPath}'. Author set to '{Environment.UserName}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
