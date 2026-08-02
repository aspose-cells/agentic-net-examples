// Title: C# – Aspose.Cells: Protect Worksheet, Allow Sorting, Block AutoFilter
// Description: Shows how to protect an Aspose.Cells worksheet in .NET, enable row sorting while disabling the AutoFilter feature, and save the workbook.
// Keywords: Aspose.Cells worksheet protection | C# allow sorting | disable AutoFilter | Protection.AllowSorting | Protection.AllowFiltering | .NET Excel security | protect Excel sheet sorting only | Aspose.Cells example | Excel worksheet lock sorting | C# Aspose.Cells protection settings
// Common Searches: Aspose.Cells protect worksheet allow sorting | C# disable AutoFilter on protected sheet Aspose.Cells | enable sorting but block filtering in Aspose.Cells .NET | WorksheetProtection AllowSorting example | Aspose.Cells set AllowFiltering false | protect Excel sheet for sorting only C#
// Developer Intent: The developer needs to secure a worksheet so users can sort data but cannot apply or change AutoFilter criteria.
// Use Cases: Distribute a reporting template where analysts may reorder rows for analysis while the filter settings stay locked. | Provide an export file that requires sorting for downstream processing but must keep filtering disabled to preserve data integrity. | Create a shared workbook where end‑users can organize data visually without altering predefined filter views.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, enable sorting, and disable AutoFilter. | Explain the effect of Protection.AllowSorting and Protection.AllowFiltering on a protected Excel sheet in Aspose.Cells for .NET. | Give a step‑by‑step tutorial for configuring worksheet protection so only sorting is permitted while all other actions are blocked.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to protect an Aspose.Cells worksheet in .NET, enable row sorting while disabling the AutoFilter feature, and save the workbook.
    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with all protection types
                sheet.Protect(ProtectionType.All);

                // Configure protection options:
                // Allow sorting of data
                sheet.Protection.AllowSorting = true;
                // Disallow use of AutoFilter (blocking filtering)
                sheet.Protection.AllowFiltering = false;

                // Save the workbook to verify the protection settings
                workbook.Save("WorksheetProtection_SortAllowed_FilterBlocked.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
