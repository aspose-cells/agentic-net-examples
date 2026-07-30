// Title: C# – Detect Worksheet Password Protection with Aspose.Cells before Editing
// Description: Shows how to use Aspose.Cells for .NET to query a worksheet’s password‑protected status via the Protection.IsProtectedWithPassword property, conditionally write data, and save the workbook.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | IsProtectedWithPassword | detect protected sheet | prevent edit on protected worksheet | protect worksheet programmatically | check sheet protection status | Aspose.Cells API
// Common Searches: How to check if an Excel worksheet is password protected using Aspose.Cells C# | Aspose.Cells C# IsProtectedWithPassword example | Prevent writing to a protected sheet with Aspose.Cells | Determine worksheet protection status before modifying cells | C# code to skip edits on password‑protected worksheet
// Developer Intent: Find out whether a worksheet is password protected and only perform data modifications when it is not.
// Use Cases: Validate protection before writing values to avoid runtime errors. | Iterate through all worksheets and update only those without a password. | Generate a report of protected versus unprotected sheets in a workbook. | Log a warning or notify users when a sheet is locked with a password.
// AI Prompts: Write C# code using Aspose.Cells to check a worksheet’s password protection and add data only if the sheet is unprotected. | Show how to remove a password from a protected worksheet after confirming its status with Aspose.Cells. | Create a C# routine that scans a workbook, lists worksheets that have password protection, and logs their names.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to use Aspose.Cells for .NET to query a worksheet’s password‑protected status via the Protection.IsProtectedWithPassword property, conditionally write data, and save the workbook.
    public class WorksheetProtectionCheckDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with a password for demonstration
                sheet.Protect(ProtectionType.All, "mySecret", null);

                // Check if the worksheet is protected with a password
                bool isProtectedWithPassword = sheet.Protection.IsProtectedWithPassword;
                Console.WriteLine($"Worksheet protected with password: {isProtectedWithPassword}");

                // Modify a cell only if the worksheet is not password protected
                if (!isProtectedWithPassword)
                {
                    sheet.Cells["A1"].PutValue("Data added");
                    Console.WriteLine("Data modification performed.");
                }
                else
                {
                    Console.WriteLine("Worksheet is password protected; modification skipped.");
                }

                // Save the workbook
                workbook.Save("WorksheetProtectionCheckDemo.xlsx");
                Console.WriteLine("Workbook saved as WorksheetProtectionCheckDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionCheckDemo.Run();
        }
    }
}
