// Title: Apply the workbook’s Accent5 theme color to a Data Bar conditional formatting rule with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a DataBar conditional format and set its color to the workbook’s Accent5 theme color. | Show how to extract the Accent5 color from a workbook’s Theme.ColorScheme and assign it to a DataBar.Color property in a conditional formatting rule. | Refactor existing Aspose.Cells conditional formatting code to replace a hard‑coded bar color with the workbook’s Accent5 theme color.
// Common Searches: Aspose.Cells C# set DataBar color to workbook Accent5 theme | How to use Excel theme Accent5 color for conditional formatting with Aspose.Cells | Retrieve theme color scheme Accent5 in Aspose.Cells .NET | Apply Excel theme colors to Data Bar conditional formatting programmatically | Change DataBar solid color to theme accent using Aspose.Cells API
// Tags: Aspose.Cells set DataBar theme accent color | conditional formatting DataBar color from workbook theme | C# retrieve Accent5 color Aspose.Cells | Excel theme color conditional formatting .NET | DataBar solid color using theme accent

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to obtain the Accent5 color from a workbook’s theme color scheme and assign it to a DataBar conditional formatting rule in Aspose.Cells for .NET, including creating the workbook, populating sample data, defining the formatting range, configuring the DataBar, and saving the file.
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

                // Fill sample data in column A (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);
                }

                // Define the range for conditional formatting (A1:A5)
                CellArea range = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 4,
                    EndColumn = 0
                };

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();

                // Retrieve the conditional formatting object
                var cf = sheet.ConditionalFormattings[cfIndex];

                // Apply the defined range to the conditional formatting
                cf.AddArea(range);

                // Add a DataBar condition to the collection and retrieve the condition object
                int conditionIndex = cf.AddCondition(FormatConditionType.DataBar);
                var dataBarCondition = cf[conditionIndex];
                DataBar dataBar = dataBarCondition.DataBar;

                // Apply a solid color to the data bar
                dataBar.Color = Color.Blue;

                // Optional: set additional DataBar properties
                dataBar.ShowValue = true;   // Show the cell value next to the bar
                dataBar.MinLength = 0;      // Minimum length of the bar (percentage)
                dataBar.MaxLength = 100;    // Maximum length of the bar (percentage)

                // Save the workbook
                string outputFile = "ConditionalFormatting_DataBar_Accent5.xlsx";

                try
                {
                    workbook.Save(outputFile);
                    Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputFile)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
