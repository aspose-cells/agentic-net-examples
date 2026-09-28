// Title: Sort Excel worksheets by creation date and reset TabId order using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, reads a 'CreatedDate' custom property from each worksheet (or uses the workbook's BuiltInDocumentProperties.CreatedTime when missing), sorts the sheets by that date, and assigns sequential TabId values starting at 0. | Write a C# program that iterates through all worksheets in a workbook, extracts their creation timestamps, orders the worksheets chronologically, updates each sheet's TabId to reflect the new order, and saves the modified workbook. | Create a script using Aspose.Cells for .NET that handles missing or invalid 'CreatedDate' properties by falling back to the workbook's creation time, then reorders the worksheet tabs accordingly and persists the changes.
// Common Searches: asp.net sort Excel sheets by custom CreatedDate property Aspose.Cells | c# reorder worksheet tabs based on creation timestamp | how to update TabId after sorting worksheets in Aspose.Cells | fallback to workbook CreatedTime when worksheet custom property not set Aspose.Cells | chronological ordering of worksheets in .xlsx using Aspose.Cells C#
// Tags: sort worksheets by creation date Aspose.Cells | update worksheet TabId .NET | read custom worksheet property CreatedDate | fallback to workbook BuiltInDocumentProperties.CreatedTime | chronological worksheet ordering C# | reassign TabId after sorting Excel sheets

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel workbook, determines a creation date for each worksheet from a custom 'CreatedDate' property or the workbook's built‑in CreatedTime, sorts the worksheets chronologically, reassigns zero‑based TabId values to match the new order, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold each worksheet together with its determined creation date
            List<(Worksheet sheet, DateTime created)> sheetInfo = new List<(Worksheet, DateTime)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                DateTime createdDate;

                try
                {
                    // Try to get a custom property named "CreatedDate"
                    var customProp = sheet.CustomProperties["CreatedDate"];
                    if (customProp != null)
                    {
                        // Retrieve the property value and parse it
                        string dateStr = customProp.Value?.ToString() ?? string.Empty;
                        if (!DateTime.TryParse(dateStr, out createdDate))
                        {
                            // Use a minimal date if parsing fails so the sheet stays at the beginning
                            createdDate = DateTime.MinValue;
                        }
                    }
                    else
                    {
                        // Fallback to the workbook's built‑in creation time
                        createdDate = workbook.BuiltInDocumentProperties.CreatedTime;
                    }
                }
                catch (Exception ex)
                {
                    // In case of any unexpected error, default to minimal date
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                    createdDate = DateTime.MinValue;
                }

                sheetInfo.Add((sheet, createdDate));
            }

            // Sort worksheets by the extracted creation dates (chronological order)
            var sortedSheets = sheetInfo.OrderBy(info => info.created).ToList();

            // Reassign TabId values based on the sorted order (zero‑based)
            for (int i = 0; i < sortedSheets.Count; i++)
            {
                sortedSheets[i].sheet.TabId = i;
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
