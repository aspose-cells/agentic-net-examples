// Title: Backup selected worksheets with a name prefix to a new Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a source workbook, filters worksheets whose names start with a given prefix, copies those sheets into a new workbook, and saves it as a backup file using Aspose.Cells. | Write a C# method that receives a source path, a destination path, and a sheet name prefix, then creates a backup workbook containing only the matching worksheets via Worksheets.AddCopy. | Provide a C# example that verifies the source Excel file exists, iterates over its worksheets, copies each sheet whose name matches a specific prefix to a new workbook, and handles any runtime errors with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy worksheets that start with 'Report_' to another workbook | How to create a backup Excel file containing only sheets with a certain prefix using Aspose.Cells | C# filter worksheets by name prefix and export to new workbook Aspose.Cells | Selective worksheet copy in Aspose.Cells .NET based on sheet name pattern
// Tags: prefix-based worksheet selection Aspose.Cells | duplicate filtered sheets into separate workbook C# | using Worksheets.AddCopy for selective export | create Excel backup with specific sheet names .NET | conditional sheet copying Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The program loads a source Excel workbook, removes the default sheet from a new workbook, iterates through all worksheets, and copies each sheet whose name begins with a defined prefix (e.g., "Report_") into the new workbook using Worksheets.AddCopy. It then saves the new workbook as a backup file, handling missing source files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Paths for source and backup workbooks
        string sourcePath = "source.xlsx";
        string backupPath = "backup.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the original workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for the backup and remove the default sheet
            Workbook backupWorkbook = new Workbook();
            backupWorkbook.Worksheets.Clear();

            // Define the prefix to filter worksheets (e.g., all sheets starting with "Report_")
            string prefix = "Report_";

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Check if the worksheet name starts with the specified prefix
                if (sheet.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    // Copy the matching worksheet into the backup workbook using its name
                    backupWorkbook.Worksheets.AddCopy(sheet.Name);
                }
            }

            // Save the backup workbook to the specified file
            backupWorkbook.Save(backupPath);
            Console.WriteLine($"Backup created successfully at: {backupPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
