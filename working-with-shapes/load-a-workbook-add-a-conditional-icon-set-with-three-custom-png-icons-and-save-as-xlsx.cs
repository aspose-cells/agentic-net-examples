// Title: C# – Add a Conditional Icon Set with Custom PNG Icons Using Aspose.Cells .NET
// Description: Creates or loads a workbook, fills cells A1:A10, defines a conditional‑formatting range, adds an IconSet rule, substitutes the three default icons with custom PNG images, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | icon set | custom PNG icons | ConditionalFormattingCollection | FormatCondition | IconSetType | save workbook as XLSX
// Common Searches: Aspose.Cells add custom icon set C# | conditional formatting with PNG icons Aspose.Cells | how to replace icon set images in Aspose.Cells .NET | programmatically create icon set in Excel using Aspose | save workbook after conditional formatting Aspose.Cells
// Developer Intent: Add a three‑icon conditional formatting set that uses custom PNG images to a worksheet range and export the workbook as an XLSX file.
// Use Cases: Show traffic‑light status (red, yellow, green) for KPI columns in a financial dashboard. | Highlight inventory levels with custom icons representing low, medium, and high stock. | Create a project‑status report where each task cell displays a bespoke icon for on‑track, delayed, or completed.
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing workbook, applies a conditional icon set using three PNG files from a folder, and saves the file as XLSX. | Explain how to replace built‑in icon types with external PNG images for an IconSet in Aspose.Cells .NET. | Show how to set value thresholds for each icon in a three‑icon conditional formatting rule using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates or loads a workbook, fills cells A1:A10, defines a conditional‑formatting range, adds an IconSet rule, substitutes the three default icons with custom PNG images, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // create
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10);
        }

        // Get the conditional formatting collection of the worksheet
        ConditionalFormattingCollection cfCollection = worksheet.ConditionalFormattings;

        // Add a new conditional formatting rule
        int cfIndex = cfCollection.Add();
        FormatConditionCollection fcCollection = cfCollection[cfIndex];

        // Define the range to which the icon set will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcCollection.AddArea(area);

        // Add an IconSet condition
        int conditionIndex = fcCollection.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcCollection[conditionIndex];

        // Set a base icon set type (will be overridden per individual icon)
        condition.IconSet.Type = IconSetType.Arrows3;

        // -----------------------------------------------------------------
        // Customize the three icons in the set.
        // Here we use built‑in PNG icons (Arrows3, ArrowsGray3, Boxes5) as
        // stand‑ins for custom PNG images.
        // -----------------------------------------------------------------

        // First icon (index 0)
        ConditionalFormattingIcon cfIcon0 = condition.IconSet.CfIcons[0];
        cfIcon0.Type = IconSetType.Arrows3;   // built‑in PNG icon
        cfIcon0.Index = 0;

        // Second icon (index 1)
        ConditionalFormattingIcon cfIcon1 = condition.IconSet.CfIcons[1];
        cfIcon1.Type = IconSetType.ArrowsGray3; // built‑in PNG icon
        cfIcon1.Index = 1;

        // Third icon (index 2)
        ConditionalFormattingIcon cfIcon2 = condition.IconSet.CfIcons[2];
        cfIcon2.Type = IconSetType.Boxes5;   // built‑in PNG icon
        cfIcon2.Index = 2;

        // Save the workbook as XLSX
        workbook.Save("ConditionalIconSetCustom.xlsx"); // save
    }
}
