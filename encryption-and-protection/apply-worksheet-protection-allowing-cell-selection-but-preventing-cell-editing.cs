// Title: Aspose.Cells .NET: Protect Worksheet, Allow Cell Selection, Block Editing (C#)
// Description: C# example that creates a workbook, protects the first worksheet with a password, enables selection of both locked and unlocked cells, disables content editing, and saves the file as WorksheetProtected.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells worksheet protection C# | protect worksheet allow selection Aspose.Cells | Aspose.Cells prevent editing locked cells | Aspose.Cells set password protection | C# Aspose.Cells Protect method | ProtectionType.All Aspose.Cells | AllowSelectingLockedCell Aspose.Cells | AllowSelectingUnlockedCell Aspose.Cells | read‑only Excel report Aspose.Cells | Excel template lock cells C#
// Common Searches: Aspose.Cells protect worksheet but still allow cell selection | C# Aspose.Cells disable editing while enabling selection | How to set a password for worksheet protection in Aspose.Cells | Allow selecting locked cells Aspose.Cells .NET | Prevent content changes in Aspose.Cells worksheet
// Developer Intent: Apply worksheet protection that lets users select any cell but prevents them from editing the sheet.
// Use Cases: Distribute a read‑only financial report where recipients can copy data but cannot modify the workbook. | Provide a template that locks all cells except designated input fields while still allowing navigation of locked cells. | Secure confidential spreadsheets with a password, permitting viewers to select and view cells without making changes.
// AI Prompts: Generate C# code with Aspose.Cells to protect a worksheet, enable selection of locked and unlocked cells, and set a password. | Explain how to configure worksheet protection in Aspose.Cells so editing is blocked but cell selection remains active. | Show how to use ProtectionType.All together with AllowSelectingLockedCell and AllowSelectingUnlockedCell flags in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // C# example that creates a workbook, protects the first worksheet with a password, enables selection of both locked and unlocked cells, disables content editing, and saves the file as WorksheetProtected.xlsx using Aspose.Cells.
    public class WorksheetProtectionExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the protection settings of the worksheet
                Protection protection = worksheet.Protection;

                // Prevent editing of locked cells
                protection.AllowEditingContent = false;

                // Allow users to select both locked and unlocked cells
                protection.AllowSelectingLockedCell = true;
                protection.AllowSelectingUnlockedCell = true;

                // Set a password for the protection
                protection.Password = "securePassword";

                // Apply protection to the worksheet (all protection types)
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("WorksheetProtected.xlsx");
                Console.WriteLine("Workbook saved as WorksheetProtected.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionExample.Run();
        }
    }
}
