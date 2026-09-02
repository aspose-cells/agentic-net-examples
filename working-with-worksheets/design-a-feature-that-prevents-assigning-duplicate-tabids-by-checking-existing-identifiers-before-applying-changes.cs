// Title: Assign a unique TabId to an Excel worksheet with Aspose.Cells for .NET while checking existing TabIds
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, scans every worksheet for a custom property named "TabId", and adds or updates that property on a target sheet only if the value is not already present. | Create a method that validates a new TabId against all existing TabId custom properties in the workbook before saving, throwing a clear exception when a duplicate is found. | Write error‑handling logic for assigning a TabId that distinguishes between missing files, nonexistent worksheets, and duplicate TabId scenarios.
// Common Searches: c# aspocells check if TabId already exists before assigning to worksheet | how to enforce unique custom property values across worksheets in Aspose.Cells | prevent duplicate TabId error when updating Excel sheet with Aspose.Cells .NET | sample code to validate unique TabId in Excel workbook using Aspose.Cells
// Tags: Aspose.Cells TabId assignment | C# Excel custom property uniqueness | prevent duplicate worksheet identifiers .NET | validate TabId across worksheets Aspose | Excel workbook custom property management

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This example demonstrates using Aspose.Cells for .NET to load an Excel workbook, collect all existing "TabId" custom property values from each worksheet, verify that a new TabId is unique, and then add or update the "TabId" property on a specified sheet before saving, with detailed exception handling for missing files, duplicate IDs, and absent worksheets.
class TabIdManager
{
    // Assigns a unique TabId to a worksheet. Throws if the TabId already exists.
    public static void AssignTabId(string workbookPath, string sheetName, string newTabId)
    {
        // Ensure the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("Workbook file not found.", workbookPath);

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Gather all existing TabIds from custom properties of each worksheet
            HashSet<string> existingIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Retrieve the custom property "TabId" if it exists
                var prop = ws.CustomProperties["TabId"];
                if (prop != null && prop.Value != null)
                {
                    existingIds.Add(prop.Value.ToString());
                }
            }

            // Verify that the new TabId is not already used
            if (existingIds.Contains(newTabId))
                throw new InvalidOperationException($"The TabId '{newTabId}' is already assigned to another worksheet.");

            // Locate the target worksheet
            Worksheet targetSheet = workbook.Worksheets[sheetName];
            if (targetSheet == null)
                throw new ArgumentException($"Worksheet '{sheetName}' does not exist in the workbook.");

            // Assign the new TabId as a custom property (adds or updates)
            var existingProp = targetSheet.CustomProperties["TabId"];
            if (existingProp != null)
            {
                // Update existing property value
                existingProp.Value = newTabId;
            }
            else
            {
                // Add new custom property
                targetSheet.CustomProperties.Add("TabId", newTabId);
            }

            // Save the workbook
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            // Wrap with additional context
            throw new ApplicationException("Failed to assign TabId.", ex);
        }
    }

    // Example usage
    static void Main()
    {
        string filePath = @"C:\Temp\Sample.xlsx";
        string sheet = "Sheet1";
        string newId = "Tab123";

        try
        {
            AssignTabId(filePath, sheet, newId);
            Console.WriteLine($"TabId '{newId}' assigned to worksheet '{sheet}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
