// Title: Detect encrypted Excel workbooks in a folder using Aspose.Cells for .NET (C#)
// Description: C# example that scans a given directory for .xls, .xlsx, .xlsm and .xlsb files, uses Aspose.Cells FileFormatUtil.DetectFileFormat to read the IsEncrypted flag, and logs each file name with its encryption status while handling missing folders and processing errors.
// Keywords: Aspose.Cells | FileFormatUtil | DetectFileFormat | IsEncrypted | encrypted Excel | password‑protected workbook | C# scan folder | list encrypted Excel files | Excel encryption detection .NET | bulk Excel audit
// Common Searches: How to check if an Excel file is password protected with Aspose.Cells C# | C# code to list encrypted .xlsx files in a directory | Detect encryption status of multiple Excel workbooks using Aspose | Log encrypted Excel file names in .NET | Aspose.Cells detect encrypted Excel files in a folder
// Developer Intent: Identify which Excel files in a specified folder are encrypted and output their names with a true/false flag.
// Use Cases: Generate an audit report of password‑protected workbooks before bulk data extraction. | Automatically skip encrypted files in an import or conversion pipeline. | Monitor a shared drive for compliance by listing encrypted Excel files and their locations.
// AI Prompts: Write a C# method that returns a List<string> of encrypted Excel file paths in a directory using Aspose.Cells. | Show how to extend the sample to recursively scan subfolders and export the file name and encryption status to a CSV file. | Explain how to handle Excel files that require a password when using FileFormatUtil.DetectFileFormat.

using System;
using System.IO;
using Aspose.Cells;

// C# example that scans a given directory for .xls, .xlsx, .xlsm and .xlsb files, uses Aspose.Cells FileFormatUtil.DetectFileFormat to read the IsEncrypted flag, and logs each file name with its encryption status while handling missing folders and processing errors.
class Program
{
    static void Main()
    {
        // Specify the directory to scan
        string folderPath = @"C:\Path\To\ExcelFiles";

        // Verify that the directory exists to avoid DirectoryNotFoundException
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Define Excel file extensions to look for
        string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };

        try
        {
            // Iterate over each extension and process matching files
            foreach (string ext in extensions)
            {
                foreach (string filePath in Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        // Detect file format and retrieve encryption information
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                        bool isEncrypted = formatInfo.IsEncrypted;

                        // Log file name and encryption status
                        Console.WriteLine($"{Path.GetFileName(filePath)} - Encrypted: {isEncrypted}");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors encountered while processing the file
                        Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors (e.g., access permissions)
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
