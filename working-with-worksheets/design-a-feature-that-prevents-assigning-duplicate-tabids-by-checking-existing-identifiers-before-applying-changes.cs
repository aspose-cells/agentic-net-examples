// Title: Prevent duplicate worksheet TabId assignments in Aspose.Cells for .NET (C#)
// Common Searches: C# code to verify a worksheet TabId is not already used before assigning it with Aspose.Cells | Aspose.Cells how to avoid TabId collision when updating worksheet properties | Check existing TabIds in a workbook using Aspose.Cells before setting a new TabId | Set unique TabId for a specific worksheet in an Excel file with Aspose.Cells .NET | Prevent duplicate TabId error in Aspose.Cells workbook manipulation
// Tags: worksheet tabid uniqueness validation Aspose.Cells | assign unique tabid to Excel worksheet C# | aspnet cells tabid collision prevention | excel workbook tabid conflict detection Aspose | prevent duplicate worksheet identifiers .NET

using System;
using System.IO;
using Aspose.Cells;

namespace TabIdDuplicatePrevention
{
    // The example defines a SetUniqueTabId method that iterates through all worksheets in a Workbook, skips the target sheet, and returns false if any other sheet already uses the requested TabId; otherwise it assigns the TabId and returns true. The Main routine demonstrates creating a workbook with preset TabIds, saving and reloading it, attempting to set a duplicate TabId (which fails), then setting a unique TabId (which succeeds), and finally saving the modified file.
    class Program
    {
        /// <param name="workbook">The workbook containing the worksheets.</param>
        /// <param name="targetSheetIndex">Zero‑based index of the worksheet to modify.</param>
        /// <param name="desiredTabId">The TabId you wish to assign.</param>
        /// <returns>True if the TabId was set successfully; otherwise false.</returns>
        static bool SetUniqueTabId(Workbook workbook, int targetSheetIndex, int desiredTabId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (targetSheetIndex < 0 || targetSheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(targetSheetIndex), "Invalid worksheet index.");

            // Scan existing worksheets for the desired TabId
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the worksheet we are about to modify
                if (ws.Index == targetSheetIndex) continue;

                if (ws.TabId == desiredTabId)
                {
                    Console.WriteLine(
                        $"Cannot assign TabId {desiredTabId} to worksheet \"{workbook.Worksheets[targetSheetIndex].Name}\" because it is already used by worksheet \"{ws.Name}\" (Index {ws.Index}).");
                    return false;
                }
            }

            // No conflict found – assign the TabId
            workbook.Worksheets[targetSheetIndex].TabId = desiredTabId;
            Console.WriteLine(
                $"TabId {desiredTabId} successfully assigned to worksheet \"{workbook.Worksheets[targetSheetIndex].Name}\".");
            return true;
        }

        static void Main(string[] args)
        {
            try
            {
                // Path to a temporary workbook file (will be created in the current directory)
                const string workbookPath = "DemoWorkbook.xlsx";

                // ------------------------------------------------------------
                // 1. Create a sample workbook with three worksheets and preset TabIds
                // ------------------------------------------------------------
                Workbook wb = new Workbook();

                // Rename the default worksheet and add two more
                wb.Worksheets[0].Name = "Summary";
                wb.Worksheets.Add("Data");
                wb.Worksheets.Add("Report");

                // Assign explicit TabIds for demonstration purposes
                wb.Worksheets[0].TabId = 101; // Summary
                wb.Worksheets[1].TabId = 102; // Data
                wb.Worksheets[2].TabId = 103; // Report

                // Save the workbook so we can reload it later (demonstrates persistence)
                wb.Save(workbookPath);
                Console.WriteLine($"Created workbook \"{workbookPath}\" with initial TabIds:");
                foreach (Worksheet ws in wb.Worksheets)
                {
                    Console.WriteLine($"  Sheet \"{ws.Name}\" (Index {ws.Index}) – TabId: {ws.TabId}");
                }

                // ------------------------------------------------------------
                // 2. Reload the workbook to simulate a real‑world scenario
                // ------------------------------------------------------------
                Workbook loadedWb;
                if (File.Exists(workbookPath))
                {
                    loadedWb = new Workbook(workbookPath);
                }
                else
                {
                    Console.WriteLine($"File \"{workbookPath}\" not found. Creating a new workbook.");
                    loadedWb = new Workbook();
                }

                Console.WriteLine("\nReloaded workbook. Current TabIds:");
                foreach (Worksheet ws in loadedWb.Worksheets)
                {
                    Console.WriteLine($"  Sheet \"{ws.Name}\" (Index {ws.Index}) – TabId: {ws.TabId}");
                }

                // ------------------------------------------------------------
                // 3. Attempt to set a duplicate TabId on the second worksheet ("Data")
                // ------------------------------------------------------------
                int targetIndex = 1; // "Data" worksheet
                int duplicateTabId = 101; // Already used by "Summary"

                Console.WriteLine(
                    $"\nAttempting to assign duplicate TabId {duplicateTabId} to worksheet \"{loadedWb.Worksheets[targetIndex].Name}\"...");
                bool success = SetUniqueTabId(loadedWb, targetIndex, duplicateTabId);
                Console.WriteLine($"Result: {(success ? "Success" : "Failed")}");

                // ------------------------------------------------------------
                // 4. Attempt to assign a unique TabId on the same worksheet
                // ------------------------------------------------------------
                int uniqueTabId = 110;
                Console.WriteLine(
                    $"\nAttempting to assign unique TabId {uniqueTabId} to worksheet \"{loadedWb.Worksheets[targetIndex].Name}\"...");
                success = SetUniqueTabId(loadedWb, targetIndex, uniqueTabId);
                Console.WriteLine($"Result: {(success ? "Success" : "Failed")}");

                // ------------------------------------------------------------
                // 5. Save the modified workbook and display final TabIds
                // ------------------------------------------------------------
                const string modifiedPath = "DemoWorkbook_Modified.xlsx";
                loadedWb.Save(modifiedPath);
                Console.WriteLine($"\nModified workbook saved as \"{modifiedPath}\". Final TabIds:");
                foreach (Worksheet ws in loadedWb.Worksheets)
                {
                    Console.WriteLine($"  Sheet \"{ws.Name}\" (Index {ws.Index}) – TabId: {ws.TabId}");
                }

                // Clean up (optional): delete temporary files if you do not need them
                // File.Delete(workbookPath);
                // File.Delete(modifiedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
