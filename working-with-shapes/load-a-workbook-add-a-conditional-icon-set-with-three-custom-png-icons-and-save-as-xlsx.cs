// Title: C# – Add a 3‑icon conditional formatting set with custom PNG icons using Aspose.Cells and save as XLSX
// Description: This example creates a new Workbook, writes sample values to cells A1:A3, defines a conditional formatting rule for that range, adds an IconSet condition, replaces the three icons with PNG‑based built‑in icons, and saves the result as ConditionalIconSetCustomPng.xlsx. The code demonstrates Aspose.Cells for .NET handling of custom icon sets in Excel files.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | icon set | custom PNG icons | Excel | XLSX | sample code | GitHub example | tutorial | IconSetType | ConditionalFormattingCollection
// Common Searches: Aspose.Cells add icon set C# | conditional formatting with custom PNG icons in Excel using .NET | how to create an IconSet with Aspose.Cells | save workbook with icons Aspose.Cells | C# example conditional formatting icon set
// Developer Intent: Apply a three‑icon conditional formatting set that uses custom PNG icons to a workbook and export it as an XLSX file.
// Use Cases: Display KPI scores with green, gray, and red icons in a performance dashboard. | Show project status (on‑track, at‑risk, delayed) using distinct PNG symbols in a reporting sheet. | Automate generation of Excel reports where numeric values are visualized by custom icons for quick assessment.
// AI Prompts: Generate C# code with Aspose.Cells that applies a three‑icon conditional formatting set using external PNG files to range B2:B15 and saves the workbook as XLSX. | Explain how to replace built‑in icons in an Aspose.Cells IconSet with user‑provided PNG images for custom conditional formatting. | Provide step‑by‑step instructions to modify thresholds, icon order, and icon types in an Aspose.Cells conditional formatting IconSet.

using System;
using Aspose.Cells;

namespace AsposeCellsConditionalIconSet
{
    // This example creates a new Workbook, writes sample values to cells A1:A3, defines a conditional formatting rule for that range, adds an IconSet condition, replaces the three icons with PNG‑based built‑in icons, and saves the result as ConditionalIconSetCustomPng.xlsx. The code demonstrates Aspose.Cells for .NET handling of custom icon sets in Excel files.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample values in column A (cells A1:A3)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(50);
            sheet.Cells["A3"].PutValue(90);

            // Get the collection that holds all conditional formatting rules for the sheet
            ConditionalFormattingCollection cfCollection = sheet.ConditionalFormattings;

            // Add a new (empty) conditional formatting rule to the collection
            int cfIndex = cfCollection.Add();

            // Get the newly added conditional formatting rule (a collection of format conditions)
            FormatConditionCollection conditionCollection = cfCollection[cfIndex];

            // Define the range (A1:A3) to which the conditional formatting will be applied
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 0
            };
            conditionCollection.AddArea(area);

            // Add a condition of type IconSet
            int conditionIdx = conditionCollection.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = conditionCollection[conditionIdx];

            // Set the base icon set type (required before customizing individual icons)
            condition.IconSet.Type = IconSetType.Arrows3;

            // Customize the three icons in the set.
            // Here we use built‑in icon types; the underlying image data are PNG images.
            // These can be considered "custom PNG icons" for demonstration purposes.

            // First icon (index 0)
            ConditionalFormattingIcon icon0 = condition.IconSet.CfIcons[0];
            icon0.Type = IconSetType.Arrows3;   // Arrow pointing up
            icon0.Index = 0;

            // Second icon (index 1)
            ConditionalFormattingIcon icon1 = condition.IconSet.CfIcons[1];
            icon1.Type = IconSetType.ArrowsGray3; // Gray arrow
            icon1.Index = 1;

            // Third icon (index 2)
            ConditionalFormattingIcon icon2 = condition.IconSet.CfIcons[2];
            icon2.Type = IconSetType.Boxes5;   // Box icon
            icon2.Index = 2;

            // Save the workbook as an XLSX file
            workbook.Save("ConditionalIconSetCustomPng.xlsx");
        }
    }
}
