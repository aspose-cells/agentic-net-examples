// Title: Protect an Aspose.Cells Worksheet in C# – Allow Row Insertion, Block Row Deletion
// Description: C# example that creates a workbook, accesses the first worksheet, configures its Protection object to enable row insertion, disable row deletion, sets a password, applies full protection with ProtectionType.All, and saves the file as WorksheetProtection.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | allow insert rows Aspose.Cells | prevent delete rows Aspose.Cells | set worksheet password Aspose.Cells | ProtectionType.All example | C# protect Excel worksheet Aspose.Cells | row insertion permission Aspose.Cells | Excel sheet protection API
// Common Searches: How to allow inserting rows while protecting a worksheet with Aspose.Cells .NET | Aspose.Cells code to stop row deletion on a protected sheet | Set password for worksheet protection using Aspose.Cells C# | Protect entire worksheet but enable specific actions in Aspose.Cells | Aspose.Cells row permission settings example
// Developer Intent: The developer needs to protect a worksheet, permit row insertion, and prevent row deletion using Aspose.Cells for .NET.
// Use Cases: Distribute a template where users can add new data rows but cannot remove existing records. | Generate a report that must stay unchanged except for appending summary rows by end users. | Provide a shared spreadsheet with locked content while allowing collaborators to insert comment rows.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, enables row insertion, disables row deletion, and applies a password. | Explain the effect of ProtectionType.All in Aspose.Cells and how to customize allowed actions such as inserting or deleting rows. | Provide a step‑by‑step tutorial for protecting an Excel worksheet in Aspose.Cells while allowing only specific operations like row insertion.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, accesses the first worksheet, configures its Protection object to enable row insertion, disable row deletion, sets a password, applies full protection with ProtectionType.All, and saves the file as WorksheetProtection.xlsx.
    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Access the worksheet's protection settings
                Protection protection = sheet.Protection;

                // Allow inserting rows while protecting the sheet
                protection.AllowInsertingRow = true;

                // Disallow deleting rows while protecting the sheet
                protection.AllowDeletingRow = false;

                // Optional: set a password for the protection
                protection.Password = "pwd123";

                // Apply protection to the worksheet (protect all aspects)
                sheet.Protect(ProtectionType.All);

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "WorksheetProtection.xlsx");

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during worksheet protection demo: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionDemo.Run();
        }
    }
}
