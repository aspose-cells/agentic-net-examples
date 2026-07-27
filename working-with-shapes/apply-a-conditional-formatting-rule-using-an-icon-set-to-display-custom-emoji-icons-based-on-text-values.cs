// Title: C# – Apply Emoji Icon Set Conditional Formatting with Aspose.Cells
// Description: Creates a workbook, writes sentiment text (Happy, Neutral, Sad) in column A, adds a hidden numeric helper column, defines a conditional‑formatting range on the helper column, applies an IconSet (Symbols3) to display emoji‑style icons, hides the numeric values, customizes each icon level, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | icon set | emoji icons | hide column | custom icons | Symbols3 | Excel automation | sentiment icons
// Common Searches: Aspose.Cells emoji icon set C# | how to hide helper column and show only icons in Aspose.Cells | customize icon set thresholds Aspose.Cells .NET | conditional formatting with emojis using Aspose.Cells | display sentiment icons instead of text in Excel with Aspose
// Developer Intent: Generate an Excel workbook that replaces text labels with emoji‑style icons by using a hidden numeric column and an IconSet conditional‑formatting rule.
// Use Cases: Convert status text (Happy/Neutral/Sad) into visual emoji icons for concise reporting. | Build dashboards where numeric scores are concealed and only emoticon icons are visible for quick assessment. | Create printable spreadsheets that convey sentiment through custom icons without exposing underlying values.
// AI Prompts: Write C# code with Aspose.Cells to map numeric thresholds to emoji icons using an IconSet and hide the helper column. | Explain how to customize individual icons in an Aspose.Cells IconSet condition for happy, neutral, and sad sentiment levels. | Provide step‑by‑step instructions to export an XLSX file that shows only emoji icons while keeping the numeric helper column invisible.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook, writes sentiment text (Happy, Neutral, Sad) in column A, adds a hidden numeric helper column, defines a conditional‑formatting range on the helper column, applies an IconSet (Symbols3) to display emoji‑style icons, hides the numeric values, customizes each icon level, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample text values that we want to represent with emojis
                sheet.Cells["A1"].PutValue("Happy");
                sheet.Cells["A2"].PutValue("Neutral");
                sheet.Cells["A3"].PutValue("Sad");

                // Add a numeric helper column that will drive the icon set thresholds
                sheet.Cells["B1"].PutValue(3); // Happy → highest icon
                sheet.Cells["B2"].PutValue(2); // Neutral → middle icon
                sheet.Cells["B3"].PutValue(1); // Sad → lowest icon

                // Hide the helper column so only the icons are visible
                sheet.Cells.HideColumn(1);

                // Define the range for the conditional formatting (the helper column)
                CellArea formatArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 2,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a new conditional formatting collection
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Apply the range
                fcc.AddArea(formatArea);

                // Add an IconSet condition
                int iconConditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
                FormatCondition iconCondition = fcc[iconConditionIdx];

                // Choose an icon set that resembles emojis (e.g., Symbols3)
                iconCondition.IconSet.Type = IconSetType.Symbols3;

                // Hide the numeric values; only icons will be shown
                iconCondition.IconSet.ShowValue = false;

                // Optionally reverse the order if you want the highest value to show the first icon
                iconCondition.IconSet.Reverse = false;

                // Customize individual icons if you want different symbols for each level
                //   Index 0 (lowest) → Red cross (as a "sad" symbol)
                //   Index 1 (middle) → Yellow exclamation (as a "neutral" symbol)
                //   Index 2 (highest) → Green check (as a "happy" symbol)
                ConditionalFormattingIcon lowIcon = iconCondition.IconSet.CfIcons[0];
                lowIcon.Type = IconSetType.Symbols3;
                lowIcon.Index = 0;

                ConditionalFormattingIcon midIcon = iconCondition.IconSet.CfIcons[1];
                midIcon.Type = IconSetType.Symbols3;
                midIcon.Index = 1;

                ConditionalFormattingIcon highIcon = iconCondition.IconSet.CfIcons[2];
                highIcon.Type = IconSetType.Symbols3;
                highIcon.Index = 2;

                // Prepare output path and ensure directory exists
                string outputPath = "EmojiIconSetConditionalFormatting.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
