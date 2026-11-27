// Title: C# example: Assign incremental TabId to new worksheets with Aspose.Cells when adding sheets
// Common Searches: how to assign incremental TabId to worksheets using Aspose.Cells in C# | Aspose.Cells set custom TabId for newly added worksheet | C# get next available TabId in an Excel workbook with Aspose.Cells | auto increment worksheet TabId when adding sheets Aspose.Cells .NET | determine maximum TabId in workbook before adding new sheet Aspose.Cells
// Tags: Aspose.Cells set worksheet TabId programmatically | C# calculate next TabId in Excel workbook | custom TabId for new sheets Aspose.Cells | worksheet TabId management Aspose.Cells | Excel TabId auto increment C#

using System;
using System.IO;
using Aspose.Cells;

namespace TabIdAssignmentDemo
{
    // The sample loads an existing workbook (or creates a new one), lists each worksheet's current TabId, adds three new worksheets, and assigns each a TabId that is one higher than the highest existing TabId using a GetNextTabId helper. The workbook is saved, reloaded, and the TabIds are printed again to confirm the incremental assignment.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to an existing workbook (if any). Change as needed.
                string inputPath = "input.xlsx";
                string outputPath = "output_with_tabids.xlsx";

                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook \"{inputPath}\" with {workbook.Worksheets.Count} worksheet(s).");
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    Console.WriteLine("Created a new workbook.");
                }

                // Display current TabIds
                Console.WriteLine("Current worksheets and their TabIds:");
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Console.WriteLine($" - {ws.Name} : TabId = {ws.TabId}");
                }

                // Add three new worksheets and assign incremental TabIds
                for (int i = 1; i <= 3; i++)
                {
                    int newIndex = workbook.Worksheets.Add();
                    Worksheet newSheet = workbook.Worksheets[newIndex];
                    newSheet.Name = $"NewSheet{i}";

                    // Determine the next TabId based on existing worksheets
                    int nextTabId = GetNextTabId(workbook);
                    newSheet.TabId = nextTabId;

                    Console.WriteLine($"Added worksheet \"{newSheet.Name}\" with assigned TabId = {newSheet.TabId}");
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");

                // Verify TabIds after saving by reloading
                Workbook verifyWorkbook = new Workbook(outputPath);
                Console.WriteLine("Verification after reload:");
                foreach (Worksheet ws in verifyWorkbook.Worksheets)
                {
                    Console.WriteLine($" - {ws.Name} : TabId = {ws.TabId}");
                }
            }
            catch (Exception ex)
            {
                // Do not hide failures; output full exception details
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Calculates the next TabId value by finding the maximum TabId among all worksheets
        /// in the given workbook and adding one.
        /// </summary>
        /// <param name="workbook">The workbook to inspect.</param>
        /// <returns>The next incremental TabId.</returns>
        private static int GetNextTabId(Workbook workbook)
        {
            int maxTabId = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // TabId can be zero if not set; treat zero as a valid value.
                if (ws.TabId > maxTabId)
                {
                    maxTabId = ws.TabId;
                }
            }

            // Increment to obtain the next identifier.
            // If all TabIds are zero, the first new sheet will receive TabId = 1.
            return maxTabId + 1;
        }
    }
}
