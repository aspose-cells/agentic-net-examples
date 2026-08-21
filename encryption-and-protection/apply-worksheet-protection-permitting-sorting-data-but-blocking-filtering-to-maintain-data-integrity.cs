// Title: Aspose.Cells C# – Protect Worksheet, Enable Sorting, Disable Filtering
// Description: Demonstrates how to protect an Aspose.Cells worksheet, allow users to sort data, and block filtering. The example creates a workbook, applies full protection, sets AllowSorting = true and AllowFiltering = false, then saves the file as WorksheetProtection.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | AllowSorting true Aspose.Cells | AllowFiltering false Aspose.Cells | protect Excel sheet enable sorting | disable filter on protected worksheet | .NET Excel protection example | Aspose.Cells security settings | C# Excel workbook protection
// Common Searches: Aspose.Cells enable sorting on protected worksheet C# | How to block filtering while protecting a sheet with Aspose.Cells | Set AllowSorting and AllowFiltering in Aspose.Cells .NET | C# code to protect Excel sheet but allow sorting only | Aspose.Cells worksheet protection options
// Developer Intent: Apply worksheet protection that permits sorting operations while preventing filter changes.
// Use Cases: Distribute a protected report where analysts can reorder rows but cannot hide data with filters. | Create a template that lets end‑users sort tables for convenience while preserving the original filter configuration. | Generate an Excel file for compliance audits where sorting is allowed for review but filtering is locked to maintain data integrity.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, enable sorting, and disable filtering. | Show how to modify an existing Aspose.Cells workbook to change AllowSorting and AllowFiltering settings. | Provide a complete .NET example that saves a protected Excel file with custom protection options.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect an Aspose.Cells worksheet, allow users to sort data, and block filtering. The example creates a workbook, applies full protection, sets AllowSorting = true and AllowFiltering = false, then saves the file as WorksheetProtection.xlsx.
    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with all protection types (no password)
                sheet.Protect(ProtectionType.All);

                // Access the protection settings
                Protection protection = sheet.Protection;

                // Allow sorting on the protected sheet
                protection.AllowSorting = true;

                // Disallow filtering on the protected sheet
                protection.AllowFiltering = false;

                // Save the workbook with the applied protection settings
                workbook.Save("WorksheetProtection.xlsx");
                Console.WriteLine("Workbook saved successfully as WorksheetProtection.xlsx");
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
            WorksheetProtectionDemo.Run();
        }
    }
}
