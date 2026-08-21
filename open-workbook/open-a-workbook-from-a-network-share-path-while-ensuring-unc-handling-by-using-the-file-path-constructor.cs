// Title: Open an Excel workbook from a UNC network share using Aspose.Cells in C#
// Description: Demonstrates how to load an Excel file from a UNC path (e.g., \\Server\Share\Folder\file.xlsx) with Aspose.Cells, create a new workbook when the file is missing, display the first worksheet name, and save a local copy. Includes basic error handling for network‑share access.
// Keywords: Aspose.Cells UNC path | C# load Excel from network share | open workbook from \Server\Share | create workbook if file not found Aspose | save workbook locally C# | Aspose.Cells file existence check | network share Excel handling .NET
// Common Searches: Aspose.Cells load workbook from UNC path | C# open Excel file on network share | How to handle missing Excel file with Aspose.Cells | Save Aspose.Cells workbook to local folder | UNC path handling in Aspose.Cells examples
// Developer Intent: Load an Excel workbook from a UNC network share, generate a new workbook when the file does not exist, and write a copy to a local directory using Aspose.Cells for .NET.
// Use Cases: Read a shared financial template from \\Server\Share, modify the first sheet, and back it up locally before further processing. | Automate nightly synchronization of a reporting workbook stored on a network drive to a local analysis folder. | Validate the presence of a master spreadsheet on a file server; if absent, create a default workbook and store it locally for later upload.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an Excel file from a UNC path, creates a new workbook when the file is missing, and saves a copy to the application's current directory. | Explain best practices for handling permissions and path formatting when loading workbooks from network shares with Aspose.Cells. | Provide a snippet that iterates all worksheets of a workbook loaded from a UNC share and logs each sheet name with its index.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an Excel file from a UNC path (e.g., \\Server\Share\Folder\file.xlsx) with Aspose.Cells, create a new workbook when the file is missing, display the first worksheet name, and save a local copy. Includes basic error handling for network‑share access.
class Program
{
    static void Main()
    {
        // UNC path to the workbook on a network share
        string uncPath = @"\\Server\Share\Folder\example.xlsx";

        Workbook workbook = null;

        try
        {
            // Load workbook if the file exists; otherwise create a new one
            if (File.Exists(uncPath))
            {
                workbook = new Workbook(uncPath);
                Console.WriteLine("Workbook loaded from UNC path.");
            }
            else
            {
                Console.WriteLine("UNC file not found. Creating a new workbook.");
                workbook = new Workbook(); // creates a new empty workbook
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Display the name of the first worksheet
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

            // Save a local copy to verify the workbook was loaded/created successfully
            string localCopyPath = Path.Combine(Environment.CurrentDirectory, "example_copy.xlsx");
            workbook.Save(localCopyPath);
            Console.WriteLine("Workbook saved locally to: " + localCopyPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
