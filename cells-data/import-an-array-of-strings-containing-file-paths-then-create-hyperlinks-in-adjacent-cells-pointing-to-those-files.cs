// Title: Create an Excel workbook with file‑path text in column A and clickable hyperlinks in column B using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through a string array of file paths, writes each path into column A of a new workbook, verifies the file exists, and inserts an "Open File" hyperlink in column B using Aspose.Cells. | Demonstrate how to convert a Windows file path to a file:// URI and use the Aspose.Cells Hyperlinks.Add method to add the link, while writing "File not found" when the target file is missing.
// Common Searches: How to create Excel hyperlinks that open local PDF files using Aspose.Cells in C# | C# Aspose.Cells add hyperlink only when the referenced file exists | Convert Windows file path to URI for Excel hyperlink with Aspose.Cells | Write an array of file paths to column A and link them in column B using Aspose.Cells
// Tags: Aspose.Cells hyperlink cell API | C# generate Excel links from local files | file existence check before Excel hyperlink Aspose | Windows path to file URI conversion C# | write file paths to worksheet column C#

using System;
using System.IO;
using Aspose.Cells;

// The example iterates over a string array of file paths, writes each path to column A of a new workbook, checks if the file exists, adds a clickable "Open File" hyperlink in column B (or a "File not found" note), and saves the result as Hyperlinks.xlsx.
class HyperlinkExample
{
    static void Main()
    {
        try
        {
            // Array of file paths to be linked
            string[] filePaths = new string[]
            {
                @"C:\Docs\file1.pdf",
                @"C:\Docs\file2.pdf",
                @"C:\Docs\file3.pdf"
            };

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Loop through the file paths and add them to column A,
            // then create a hyperlink in column B pointing to each file
            for (int i = 0; i < filePaths.Length; i++)
            {
                // Write the file path text in column A (zero‑based index)
                sheet.Cells[i, 0].PutValue(filePaths[i]);

                // Verify that the target file exists to avoid FileNotFoundException
                if (File.Exists(filePaths[i]))
                {
                    // Convert Windows path to URI format for the hyperlink
                    string uri = "file:///" + filePaths[i].Replace("\\", "/");

                    // Add a hyperlink to column B (same row)
                    // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
                    sheet.Hyperlinks.Add(i, 1, 1, 1, uri);

                    // Set display text for the hyperlink cell
                    sheet.Cells[i, 1].PutValue("Open File");
                }
                else
                {
                    // If the file does not exist, indicate it in column B
                    sheet.Cells[i, 1].PutValue("File not found");
                }
            }

            // Save the workbook to a file
            workbook.Save("Hyperlinks.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
