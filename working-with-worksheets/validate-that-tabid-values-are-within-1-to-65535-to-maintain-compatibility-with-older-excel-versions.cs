// Title: Clamp Worksheet TabId to 1‑65535 with Aspose.Cells for .NET
// Description: Shows how to validate a worksheet's TabId, enforce the 1‑65535 range required for Excel 97‑2003 compatibility, assign the corrected value, save the workbook, and confirm the setting using Aspose.Cells in C#.
// Keywords: Aspose.Cells TabId | worksheet TabId range | Excel 97-2003 compatibility | C# clamp TabId | validate TabId Aspose | legacy Excel TabId limit
// Common Searches: Aspose.Cells set TabId range | How to limit worksheet TabId to 65535 | TabId out of range error Aspose | Validate TabId before saving Excel file | C# example for TabId validation
// Developer Intent: Guarantee that a worksheet's TabId remains within the 1‑65535 limits so the workbook is compatible with older Excel versions.
// Use Cases: Check user‑supplied TabId values before assigning them to avoid save‑time exceptions in legacy Excel formats. | Automatically adjust out‑of‑range TabId numbers to the nearest valid boundary when generating automated reports. | Reload a saved workbook to verify that the corrected TabId persisted correctly.
// AI Prompts: Write C# code with Aspose.Cells that verifies a TabId, clamps it to 1‑65535, applies it to a worksheet, and saves the file. | Explain Aspose.Cells' behavior when a TabId exceeds the Excel 97‑2003 limit and outline a best‑practice correction pattern. | Provide a script to batch‑process multiple worksheets, ensuring each TabId is validated and corrected before export.

using System;
using Aspose.Cells;

// Shows how to validate a worksheet's TabId, enforce the 1‑65535 range required for Excel 97‑2003 compatibility, assign the corrected value, save the workbook, and confirm the setting using Aspose.Cells in C#.
class TabIdValidator
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Desired TabId value (example). Replace with your actual value.
        int desiredTabId = 70000;

        // Validate that TabId is within the range 1 to 65535 (Excel 97-2003 compatibility)
        if (desiredTabId < 1 || desiredTabId > 65535)
        {
            Console.WriteLine($"TabId {desiredTabId} is out of the valid range (1‑65535). Adjusting to the nearest valid value.");
            // Clamp the value to the allowed range
            desiredTabId = Math.Max(1, Math.Min(65535, desiredTabId));
        }

        // Assign the validated TabId to the worksheet
        worksheet.TabId = desiredTabId;
        Console.WriteLine($"Worksheet TabId set to {worksheet.TabId}");

        // Save the workbook
        string outputPath = "TabIdValidated.xlsx";
        workbook.Save(outputPath);

        // Load the saved workbook to verify the TabId
        Workbook loadedWorkbook = new Workbook(outputPath);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
        Console.WriteLine($"Loaded Worksheet TabId: {loadedWorksheet.TabId}");
    }
}
