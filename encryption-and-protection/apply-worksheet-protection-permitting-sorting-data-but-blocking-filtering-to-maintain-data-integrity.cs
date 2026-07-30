// Title: Aspose.Cells for .NET – Protect Worksheet while Enabling Sorting and Disabling Filtering
// Description: Demonstrates how to protect an Excel worksheet using Aspose.Cells, grant users the ability to sort data, and prevent them from applying filters. The workbook is saved without a password as WorksheetProtection_SortAllowed_FilterBlocked.xlsx.
// Keywords: Aspose.Cells protect worksheet .NET | allow sorting on protected sheet | disable filtering Excel protection | worksheet protection options | C# Aspose.Cells example
// Common Searches: protect Excel sheet but keep sorting enabled Aspose.Cells | how to block filters on a protected worksheet using .NET | Aspose.Cells allow sort disallow filter | worksheet protection without password C#
// Developer Intent: Apply worksheet protection that permits sorting operations while restricting filter usage via Aspose.Cells for .NET.
// Use Cases: Financial dashboards where analysts can reorder rows but must view a consistent filter set. | Shared data templates that allow sorting for quick insights yet lock filter criteria to preserve data integrity. | Reporting tools that need sortable tables without exposing filter controls to end‑users.
// AI Prompts: Generate C# code with Aspose.Cells to protect a worksheet, enable sorting, and disable filtering. | Explain the difference between AllowSorting and AllowFiltering properties in Aspose.Cells worksheet protection. | Show how to apply distinct protection settings to multiple worksheets in a single workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Entry point for the application
    // Demonstrates how to protect an Excel worksheet using Aspose.Cells, grant users the ability to sort data, and prevent them from applying filters. The workbook is saved without a password as WorksheetProtection_SortAllowed_FilterBlocked.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                WorksheetProtectionDemo.Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with all protection types (no password)
            worksheet.Protect(ProtectionType.All);

            // Allow sorting on the protected worksheet
            worksheet.Protection.AllowSorting = true;

            // Disallow filtering on the protected worksheet
            worksheet.Protection.AllowFiltering = false;

            // Save the workbook to verify the protection settings
            workbook.Save("WorksheetProtection_SortAllowed_FilterBlocked.xlsx");
        }
    }
}
