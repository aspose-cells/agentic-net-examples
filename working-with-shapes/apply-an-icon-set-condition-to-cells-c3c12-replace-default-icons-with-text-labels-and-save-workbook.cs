// Title: Aspose.Cells for .NET – Apply IconSet Conditional Formatting to C3:C12 and display text labels only
// Description: C# example that creates a workbook, fills cells C3‑C12 with numbers, adds an IconSet conditional format, hides the icons by setting IconSet.Type to None, shows the cell values as text, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells IconSet C# | hide icons conditional formatting | show values instead of icons | IconSet Type None | Excel conditional formatting programmatic | save workbook Aspose.Cells
// Common Searches: Aspose.Cells hide IconSet icons | C# replace Excel conditional icons with text | IconSet show value only Aspose.Cells | apply IconSet to range C3:C12 .NET | save workbook after conditional formatting Aspose
// Developer Intent: Generate a workbook, apply an IconSet rule to C3:C12, suppress the icons, display the numeric values as text, and write the file to disk.
// Use Cases: Create accessibility‑compliant reports that use text instead of visual icons. | Export data to systems that cannot interpret Excel icon graphics. | Demonstrate how to modify IconSet properties (Type, ShowValue) via Aspose.Cells API.
// AI Prompts: Write C# code with Aspose.Cells to add an IconSet conditional format to a range and configure it to hide icons while showing cell values. | Explain the effect of setting IconSet.Type = None and ShowValue = true in Aspose.Cells for .NET. | Provide verification steps to ensure the saved workbook displays only numeric labels in the specified cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, fills cells C3‑C12 with numbers, adds an IconSet conditional format, hides the icons by setting IconSet.Type to None, shows the cell values as text, and saves the file as an .xlsx document.
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

                // Populate cells C3:C12 with sample numeric values
                // (rows are zero‑based, column C is index 2)
                for (int row = 2; row <= 11; row++)
                {
                    sheet.Cells[row, 2].PutValue((row - 1) * 10); // 10, 20, …, 110
                }

                // Add an empty conditional formatting collection
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

                // Define the range C3:C12
                CellArea area = new CellArea
                {
                    StartRow = 2,   // C3
                    EndRow = 11,    // C12
                    StartColumn = 2,
                    EndColumn = 2
                };
                fcs.AddArea(area);

                // Add an IconSet condition
                int conditionIdx = fcs.AddCondition(FormatConditionType.IconSet);
                FormatCondition condition = fcs[conditionIdx];

                // Configure the IconSet to hide icons and show cell values as text
                IconSet iconSet = condition.IconSet;
                iconSet.Type = IconSetType.None;   // Remove default icons
                iconSet.ShowValue = true;          // Show the cell values (text labels)
                iconSet.Reverse = false;           // Keep default order

                // Define output file name
                string outputPath = "IconSetWithTextLabels.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
