using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    public class WorksheetSelectionProtection
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the worksheet protection settings
                Protection protection = worksheet.Protection;

                // Allow users to select both locked and unlocked cells
                protection.AllowSelectingLockedCell = true;
                protection.AllowSelectingUnlockedCell = true;

                // Prevent editing of locked cells (default is false, set explicitly for clarity)
                protection.AllowEditingContent = false;

                // Set a password for the protection (optional but common)
                protection.Password = "pwd123";

                // Apply protection to the worksheet (protect all aspects)
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("ProtectedSelectionOnly.xlsx");
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
            WorksheetSelectionProtection.Run();
        }
    }
}