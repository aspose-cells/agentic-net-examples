// Title: Open Excel workbook from a UNC network share with Aspose.Cells (.NET) and save locally
// Description: Shows how to verify a file on a Windows UNC share, load it using Aspose.Cells' Workbook(string) constructor, fall back to a new workbook when the file is absent, display the worksheet count, ensure the target folder exists, and save the workbook to a local path with proper exception handling in C#.
// Keywords: Aspose.Cells UNC path | Aspose.Cells load workbook from network share | C# open Excel from UNC | Workbook(string) constructor | File.Exists UNC | Directory.CreateDirectory save workbook | exception handling Aspose.Cells | .NET network share Excel | save workbook to local folder | fallback workbook Aspose.Cells
// Common Searches: How to open an Excel file on a UNC share with Aspose.Cells | Aspose.Cells load workbook from network share C# | Save Aspose.Cells workbook to local folder after opening from UNC | Create workbook when file not found Aspose.Cells | Ensure directory exists before saving workbook Aspose.Cells
// Developer Intent: Load an Excel file from a UNC share, create a new workbook if the file does not exist, and optionally write a copy to a local directory.
// Use Cases: Open a workbook located on a Windows network share after confirming its presence. | Automatically generate a new workbook with a default sheet when the shared file is missing. | Save the opened or newly created workbook to a local path, creating the destination folder if it isn’t already present.
// AI Prompts: Generate C# code that opens an Excel file from a UNC path using Aspose.Cells, checks existence, and creates a new workbook if missing. | Show how to save a workbook loaded from a network share to a local directory, creating the folder if needed, with try‑catch error handling. | Explain permission and authentication considerations for accessing a Windows UNC share with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to verify a file on a Windows UNC share, load it using Aspose.Cells' Workbook(string) constructor, fall back to a new workbook when the file is absent, display the worksheet count, ensure the target folder exists, and save the workbook to a local path with proper exception handling in C#.
class Program
{
    static void Main()
    {
        // UNC path to the workbook on a network share
        string uncPath = @"\\ServerName\ShareFolder\Sample.xlsx";

        Workbook workbook = null;

        try
        {
            // Verify that the network file exists before attempting to load it
            if (File.Exists(uncPath))
            {
                // Load the workbook from the UNC path
                workbook = new Workbook(uncPath);
                Console.WriteLine("Workbook opened from: " + uncPath);
            }
            else
            {
                // If the file is not found, create a new workbook as a fallback
                Console.WriteLine("Network file not found. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Display basic information
            Console.WriteLine("Worksheet count: " + workbook.Worksheets.Count);

            // Optionally save a copy to a local folder
            string localCopyPath = @"C:\Temp\SampleCopy.xlsx";

            // Ensure the target directory exists
            string localDir = Path.GetDirectoryName(localCopyPath);
            if (!Directory.Exists(localDir))
            {
                Directory.CreateDirectory(localDir);
            }

            workbook.Save(localCopyPath);
            Console.WriteLine("Workbook saved to: " + localCopyPath);
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the error
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
