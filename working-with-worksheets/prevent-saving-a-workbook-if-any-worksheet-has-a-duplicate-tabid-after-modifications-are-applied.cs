// Title: How to abort saving an Aspose.Cells workbook when worksheets contain duplicate TabId values in C#
// Common Searches: C# Aspose.Cells check for duplicate worksheet TabId before saving | prevent workbook.save if TabId conflict exists Aspose.Cells .NET | detect duplicate TabId values across worksheets using Aspose.Cells API | validate unique TabId for each worksheet in Aspose.Cells C# | Aspose.Cells .NET abort save on duplicate worksheet TabId
// Tags: Aspose.Cells duplicate TabId detection | worksheet TabId uniqueness validation | abort workbook save on TabId conflict | C# Aspose.Cells TabId duplicate check | prevent saving workbook with duplicate TabId

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TabIdDuplicateCheck
{
    // The example loads or creates a workbook, applies modifications that assign TabId values to worksheets, then runs a HasDuplicateTabIds routine to identify any repeated TabId. If duplicates are found, the program writes an error and exits without calling Workbook.Save; otherwise it saves the workbook to the specified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : null;
                string outputPath = args.Length > 1 ? args[1] : "Result.xlsx";

                Workbook workbook;

                if (!string.IsNullOrEmpty(inputPath))
                {
                    if (!File.Exists(inputPath))
                    {
                        Console.Error.WriteLine($"Error: Input file '{inputPath}' does not exist.");
                        Environment.Exit(1);
                        return;
                    }

                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from '{inputPath}'.");
                }
                else
                {
                    // Create a new workbook with three worksheets for demonstration
                    workbook = new Workbook();
                    workbook.Worksheets.Add("Sheet2");
                    workbook.Worksheets.Add("Sheet3");
                    Console.WriteLine("Created a new workbook with default worksheets.");
                }

                // -------------------------------------------------
                // Example modifications – you can replace this block
                // with any logic that changes worksheets, e.g., adding,
                // removing, renaming, or setting TabId values.
                // -------------------------------------------------
                ApplySampleModifications(workbook);
                // -------------------------------------------------

                // Validate TabId uniqueness before saving
                if (HasDuplicateTabIds(workbook, out var duplicateIds))
                {
                    Console.Error.WriteLine("Error: Duplicate TabId values detected. The workbook will NOT be saved.");
                    Console.Error.WriteLine("Duplicated TabId(s): " + string.Join(", ", duplicateIds));
                    Environment.Exit(2);
                }
                else
                {
                    // Save the workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An unexpected error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(99);
            }
        }

        /// <summary>
        /// Applies sample modifications to the workbook.
        /// This method demonstrates how a user might change worksheets
        /// and explicitly set TabId values.
        /// </summary>
        private static void ApplySampleModifications(Workbook workbook)
        {
            // Example: set TabId for each worksheet to its index + 100
            // (In a real scenario, TabId could be set based on external data)
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                int newTabId = 100 + i; // simple deterministic value
                ws.TabId = newTabId;
                Console.WriteLine($"Worksheet '{ws.Name}' assigned TabId = {newTabId}");
            }

            // Example: add a new worksheet and intentionally create a duplicate TabId
            // Comment out the following block if you do NOT want a duplicate for testing.
            /*
            Worksheet dupSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            dupSheet.Name = "DuplicateTabIdSheet";
            // Force a duplicate TabId (same as first worksheet)
            dupSheet.TabId = workbook.Worksheets[0].TabId;
            Console.WriteLine($"Added worksheet '{dupSheet.Name}' with duplicate TabId = {dupSheet.TabId}");
            */
        }

        /// <summary>
        /// Checks whether any worksheets in the workbook share the same TabId.
        /// </summary>
        /// <param name="workbook">The workbook to inspect.</param>
        /// <param name="duplicateIds">
        /// When the method returns true, this collection contains the TabId values that appear more than once.
        /// </param>
        /// <returns>True if duplicates exist; otherwise false.</returns>
        private static bool HasDuplicateTabIds(Workbook workbook, out List<int> duplicateIds)
        {
            var seen = new HashSet<int>();
            var duplicates = new HashSet<int>();

            foreach (Worksheet ws in workbook.Worksheets)
            {
                int tabId = ws.TabId;
                if (!seen.Add(tabId))
                {
                    // TabId already existed – record as duplicate
                    duplicates.Add(tabId);
                }
            }

            duplicateIds = new List<int>(duplicates);
            return duplicateIds.Count > 0;
        }
    }
}
