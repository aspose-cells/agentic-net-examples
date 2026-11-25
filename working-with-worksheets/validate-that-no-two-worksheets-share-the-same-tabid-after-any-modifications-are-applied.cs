// Title: C# code to detect and fix duplicate worksheet TabId values in an Aspose.Cells workbook
// Common Searches: how to ensure each worksheet has a unique TabId using Aspose.Cells in C# | detect duplicate TabId in an Excel workbook with Aspose.Cells .NET | resolve worksheet TabId conflicts after adding new sheets in Aspose.Cells | C# program to validate worksheet TabId uniqueness in an XLSX file | automatically reassign duplicate TabIds in an Aspose.Cells workbook
// Tags: Aspose.Cells duplicate worksheet TabId detection | C# validate unique worksheet TabId | resolve worksheet TabId conflicts Aspose.Cells | assign new TabId to Excel sheets .NET | worksheet TabId uniqueness enforcement

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetTabIdValidator
{
    // The example loads a workbook (or creates a sample one with intentional duplicate TabId values), applies modifications that may introduce further duplicates, validates that every worksheet’s TabId is unique, reports any duplicate TabIds with the affected sheet indexes, automatically resolves conflicts by assigning new incremental TabId values to later-occurring sheets, and saves the corrected workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = null;
                string outputPath = null;

                if (args.Length >= 1)
                {
                    inputPath = args[0];
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Error: Input file '{inputPath}' does not exist.");
                        return;
                    }

                    if (args.Length >= 2)
                    {
                        outputPath = args[1];
                    }
                }

                // Load or create workbook
                Workbook workbook;
                if (inputPath != null)
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from '{inputPath}'.");
                }
                else
                {
                    // Create a sample workbook with intentional duplicate TabIds
                    workbook = CreateSampleWorkbookWithDuplicates();
                    Console.WriteLine("Created sample workbook with intentional duplicate TabIds.");
                }

                // Apply any modifications you need here.
                // For demonstration, we will add a new worksheet and set its TabId to an existing value.
                ApplySampleModifications(workbook);

                // Validate uniqueness of TabIds
                bool isValid = ValidateUniqueTabIds(workbook, out List<(int TabId, List<int> SheetIndexes)> duplicates);

                if (isValid)
                {
                    Console.WriteLine("Validation passed: All worksheets have unique TabId values.");
                }
                else
                {
                    Console.WriteLine("Validation failed: Duplicate TabId values detected.");
                    foreach (var dup in duplicates)
                    {
                        Console.WriteLine($"  TabId {dup.TabId} is used by worksheets at indexes: {string.Join(", ", dup.SheetIndexes)}");
                    }

                    // Resolve duplicates by assigning new unique TabIds
                    ResolveDuplicateTabIds(workbook, duplicates);
                    Console.WriteLine("Duplicate TabIds have been resolved by assigning new unique values.");
                }

                // Save the workbook if an output path is provided or if we created a sample workbook
                if (outputPath == null)
                {
                    // If no output path, generate one based on input or temporary name
                    outputPath = inputPath != null
                        ? Path.Combine(Path.GetDirectoryName(inputPath) ?? "", Path.GetFileNameWithoutExtension(inputPath) + "_validated.xlsx")
                        : Path.Combine(Directory.GetCurrentDirectory(), "ValidatedWorkbook.xlsx");
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Creates a workbook with three worksheets where two of them share the same TabId.
        /// </summary>
        private static Workbook CreateSampleWorkbookWithDuplicates()
        {
            Workbook wb = new Workbook();

            // Ensure we have at least three sheets
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Name = "SheetOne";
            sheet1.TabId = 101;

            Worksheet sheet2 = wb.Worksheets.Add("SheetTwo");
            sheet2.TabId = 102;

            Worksheet sheet3 = wb.Worksheets.Add("SheetThree");
            // Intentional duplicate TabId (same as SheetOne)
            sheet3.TabId = 101;

            return wb;
        }

        /// <summary>
        /// Demonstrates modifications that could potentially introduce duplicate TabIds.
        /// </summary>
        private static void ApplySampleModifications(Workbook workbook)
        {
            // Add a new worksheet
            Worksheet newSheet = workbook.Worksheets.Add("NewSheet");

            // Intentionally set its TabId to an existing one to trigger validation
            // We'll use the TabId of the first worksheet
            int existingTabId = workbook.Worksheets[0].TabId;
            newSheet.TabId = existingTabId;

            Console.WriteLine($"Added worksheet '{newSheet.Name}' with TabId set to {newSheet.TabId} (duplicate of first sheet).");
        }

        /// <summary>
        /// Checks that each worksheet in the workbook has a unique TabId.
        /// Returns true if all TabIds are unique; otherwise false.
        /// </summary>
        private static bool ValidateUniqueTabIds(Workbook workbook, out List<(int TabId, List<int> SheetIndexes)> duplicateInfo)
        {
            var tabIdMap = new Dictionary<int, List<int>>(); // TabId -> list of worksheet indexes
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                int tabId = workbook.Worksheets[i].TabId;
                if (!tabIdMap.ContainsKey(tabId))
                {
                    tabIdMap[tabId] = new List<int>();
                }
                tabIdMap[tabId].Add(i);
            }

            duplicateInfo = new List<(int TabId, List<int> SheetIndexes)>();
            foreach (var kvp in tabIdMap)
            {
                if (kvp.Value.Count > 1)
                {
                    duplicateInfo.Add((kvp.Key, kvp.Value));
                }
            }

            return duplicateInfo.Count == 0;
        }

        /// <summary>
        /// Resolves duplicate TabIds by assigning new unique values to the later-occurring worksheets.
        /// </summary>
        private static void ResolveDuplicateTabIds(Workbook workbook, List<(int TabId, List<int> SheetIndexes)> duplicates)
        {
            // Determine the highest TabId currently used
            int maxTabId = int.MinValue;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                int current = workbook.Worksheets[i].TabId;
                if (current > maxTabId) maxTabId = current;
            }

            // For each duplicate group, keep the first occurrence unchanged and reassign the rest
            foreach (var dup in duplicates)
            {
                // Skip the first index (keep its TabId)
                for (int j = 1; j < dup.SheetIndexes.Count; j++)
                {
                    int sheetIdx = dup.SheetIndexes[j];
                    maxTabId++; // Increment to get a new unique TabId
                    workbook.Worksheets[sheetIdx].TabId = maxTabId;
                    Console.WriteLine($"Worksheet '{workbook.Worksheets[sheetIdx].Name}' (index {sheetIdx}) assigned new TabId {maxTabId}.");
                }
            }
        }
    }
}
