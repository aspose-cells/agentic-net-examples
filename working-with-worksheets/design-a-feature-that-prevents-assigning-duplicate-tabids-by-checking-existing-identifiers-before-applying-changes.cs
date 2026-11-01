// Title: Prevent assigning duplicate TabId custom properties to worksheets using Aspose.Cells for .NET
// AI Prompts: Write a C# method that adds a TabId custom property to a worksheet only after confirming the value does not exist on any other sheet in the workbook with Aspose.Cells. | Refactor the TryAssignTabId method to throw an InvalidOperationException when a duplicate TabId is detected instead of returning a Boolean flag. | Create a utility that scans all worksheets in a workbook and returns a list of the current TabId custom property values.
// Common Searches: aspnet aspose.cells check for existing TabId before assigning to worksheet | c# ensure unique TabId across all worksheets in an Excel file | how to validate custom property uniqueness in Aspose.Cells workbook | prevent duplicate sheet identifiers when using Aspose.Cells for .NET
// Tags: worksheet tabid uniqueness Aspose.Cells | custom property duplicate check C# | assign tabid to worksheet Aspose.Cells | sheet identifier duplication guard .NET | validate worksheet custom property uniqueness

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsTabIdManager
{
    // The TabIdHelper class loads a workbook, provides TryAssignTabId which gathers existing "TabId" custom properties from all worksheets, rejects duplicates, assigns the new TabId when unique, and includes demo and save methods.
    public class TabIdHelper
    {
        private readonly Workbook _workbook;

        // Load an existing workbook (lifecycle rule: load)
        public TabIdHelper(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Input workbook not found: {filePath}");

            try
            {
                _workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                throw;
            }
        }

        // Create a new workbook (lifecycle rule: create) - optional helper
        public TabIdHelper()
        {
            _workbook = new Workbook();
        }

        // Retrieves the TabId stored in a worksheet's custom property "TabId"
        private string GetTabId(Worksheet sheet)
        {
            var prop = sheet.CustomProperties.FirstOrDefault(p => p.Name == "TabId");
            return prop?.Value?.ToString();
        }

        // Assigns a TabId to a worksheet after ensuring it is not already used
        public bool TryAssignTabId(string sheetName, string newTabId)
        {
            // Validate input
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(newTabId))
                throw new ArgumentException("Sheet name and TabId must be non‑empty.");

            // Find the target worksheet
            Worksheet targetSheet = _workbook.Worksheets[sheetName];
            if (targetSheet == null)
                throw new ArgumentException($"Worksheet '{sheetName}' does not exist.");

            // Collect all existing TabIds (excluding the target sheet itself)
            var existingIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in _workbook.Worksheets)
            {
                if (ws == targetSheet) continue; // skip the sheet we are updating
                string id = GetTabId(ws);
                if (!string.IsNullOrEmpty(id))
                    existingIds.Add(id);
            }

            // Check for duplicate
            if (existingIds.Contains(newTabId))
                return false; // Duplicate found – do not assign

            // No duplicate – assign the TabId (store as a custom property)
            var existingProp = targetSheet.CustomProperties.FirstOrDefault(p => p.Name == "TabId");
            if (existingProp != null)
                existingProp.Value = newTabId;
            else
                targetSheet.CustomProperties.Add("TabId", newTabId);

            return true;
        }

        // Example method to demonstrate usage
        public void Demo()
        {
            try
            {
                // Attempt to set TabId for "Sheet1"
                bool success = TryAssignTabId("Sheet1", "TAB_001");
                Console.WriteLine(success
                    ? "TabId assigned successfully."
                    : "Duplicate TabId detected – assignment aborted.");

                // Attempt to set a duplicate TabId for another sheet
                success = TryAssignTabId("Sheet2", "TAB_001");
                Console.WriteLine(success
                    ? "TabId assigned successfully."
                    : "Duplicate TabId detected – assignment aborted.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Demo error: {ex.Message}");
            }
        }

        // Save the workbook (lifecycle rule: save)
        public void Save(string outputPath)
        {
            try
            {
                _workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
                throw;
            }
        }
    }

    // Sample program entry point
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook
                var helper = new TabIdHelper("InputWorkbook.xlsx");

                // Run demo to show duplicate checking
                helper.Demo();

                // Save changes
                helper.Save("OutputWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
