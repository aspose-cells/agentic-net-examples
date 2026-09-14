// Title: Copy worksheets whose names start with a prefix into a new workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing Excel file, selects all worksheets whose names begin with a given prefix, copies each to a newly created workbook, and saves the backup file using Aspose.Cells. | Modify the example to accept command‑line arguments for source path, destination path, and worksheet name prefix, then perform the same filtered copy operation.
// Common Searches: Aspose.Cells C# copy only worksheets that start with a specific prefix to another workbook | How to create a backup workbook containing selected sheets based on name pattern in .NET | Filter worksheets by name prefix and export them to a new Excel file using Aspose.Cells API
// Tags: copy worksheets by prefix Aspose.Cells | create backup workbook from filtered sheets C# | addcopy worksheet to another workbook Aspose.Cells | filter worksheet names using StartsWith Aspose.Cells | select worksheets with name pattern .NET

using Aspose.Cells;
using System;
using System.IO;

// Loads source.xlsx, copies every worksheet whose name begins with "Report_" into a new workbook, and saves it as backup.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string backupPath = "backup.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for backup and clear default sheet
            Workbook backupWorkbook = new Workbook();
            backupWorkbook.Worksheets.Clear();

            // Define the prefix to filter worksheets
            string prefix = "Report_";

            // Copy worksheets whose names start with the specified prefix
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                if (ws.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    // Add a copy of the worksheet to the backup workbook using its name
                    backupWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Save the backup workbook
            backupWorkbook.Save(backupPath);
            Console.WriteLine($"Backup saved to {backupPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
