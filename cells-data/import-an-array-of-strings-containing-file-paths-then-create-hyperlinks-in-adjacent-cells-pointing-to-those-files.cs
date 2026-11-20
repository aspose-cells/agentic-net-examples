// Title: Generate adjacent file‑path hyperlinks in Excel using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to a cell from a string array of file paths | Generate an Excel sheet with file path column and adjacent Open File hyperlink column using Aspose.Cells | How to check if a file exists before creating a hyperlink in Aspose.Cells worksheet | Programmatically create clickable links to local files in Excel with Aspose.Cells .NET | Save an Excel workbook with hyperlinks to PDFs and DOCX files using Aspose.Cells
// Tags: Aspose.Cells insert cell hyperlink | C# write file paths to Excel column | hyperlink column B referencing column A file paths | validate file existence before hyperlink creation | save workbook as XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkExample
{
    // The example reads a string array of local file paths, optionally verifies each file's existence, writes the paths to column A of a new workbook, places the text "Open File" in column B, adds a hyperlink in column B that points to the corresponding file, and saves the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Array of file paths to be inserted into the worksheet
                string[] filePaths = new string[]
                {
                    @"C:\Docs\file1.pdf",
                    @"C:\Docs\file2.docx",
                    @"C:\Docs\file3.xlsx"
                };

                // Verify that each file exists to avoid runtime errors
                foreach (string path in filePaths)
                {
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"Warning: File not found - {path}");
                    }
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Loop through the file paths and write them to column A,
                // then create a hyperlink in column B pointing to the same file
                for (int i = 0; i < filePaths.Length; i++)
                {
                    // Write the file path text in column A (zero‑based index)
                    worksheet.Cells[i, 0].PutValue(filePaths[i]);

                    // Set display text for the hyperlink in column B
                    worksheet.Cells[i, 1].PutValue("Open File");

                    // Add a hyperlink to the cell in column B.
                    // Using the overload that specifies the range (single cell).
                    worksheet.Hyperlinks.Add(i, 1, i, 1, filePaths[i]);
                }

                // Save the workbook to an XLSX file
                string outputPath = "HyperlinksOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
