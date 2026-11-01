// Title: Set the Subject built‑in property of an Excel workbook from the most frequent value in the first column using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook Subject property based on column data | How to automatically assign Excel Subject document property from worksheet content with Aspose.Cells | Find most common string in first column and update Excel built‑in properties using Aspose.Cells | Aspose.Cells example for deriving main topic from worksheet and setting Subject property
// Tags: Aspose.Cells set workbook subject property | C# extract most frequent column value | Excel built‑in document property automation | derive main topic from worksheet data | update Excel properties programmatically Aspose

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;

// The sample loads an existing Excel file, scans the first column of the first worksheet to identify the most frequent non‑empty string, treats that string as the main topic (or defaults to "Untitled"), assigns it to the workbook's built‑in Subject property, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Analyze the first worksheet to determine the main topic
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Example analysis: find the most frequent non‑empty string in the first column
        Dictionary<string, int> frequency = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            string text = cells[row, 0].StringValue?.Trim();
            if (!string.IsNullOrEmpty(text))
            {
                if (frequency.ContainsKey(text))
                    frequency[text]++;
                else
                    frequency[text] = 1;
            }
        }

        // Determine the main topic (fallback to "Untitled" if no data)
        string mainTopic = frequency.OrderByDescending(kv => kv.Value)
                                    .Select(kv => kv.Key)
                                    .FirstOrDefault() ?? "Untitled";

        // Update the Subject built‑in property with the derived main topic
        workbook.BuiltInDocumentProperties.Subject = mainTopic;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
