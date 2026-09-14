// Title: Identify nillable ContentTypeProperty entries in an Excel sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, scans each row for a property name and an optional flag, and returns a comma‑separated list of property names where the flag is true. | Update the program to output both the list of nillable property names and the total count, avoiding hard‑coded column indexes. | Add logic to detect the header row, locate the column titled "Optional" (or a similar name) dynamically, and then generate the nillable properties summary.
// Common Searches: how to list optional content type properties from an Excel file using Aspose.Cells in C# | C# Aspose.Cells read boolean values and generate nillable property summary | extract rows with true optional flag from Excel worksheet using Aspose.Cells | generate comma separated list of nillable properties from Excel with .NET
// Tags: Aspose.Cells iterate rows and evaluate optional flag | C# generate nillable property list from Excel | parse boolean or text optional column Aspose.Cells | dynamic header column lookup Aspose.Cells | output comma‑separated nillable properties .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads 'ContentTypes.xlsx', iterates data rows, interprets the optional column (boolean or 'true'/'false'), collects property names marked as optional, and prints a summary listing the nillable properties or indicates none were found.
class Program
{
    static void Main()
    {
        const string filePath = "ContentTypes.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The required file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the ContentTypeProperty definitions
            Workbook workbook = new Workbook(filePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range to iterate over rows
            Cells cells = sheet.Cells;
            int firstRow = cells.MinRow;
            int lastRow = cells.MaxRow;

            // List to hold the names of properties that are marked as optional (nillable)
            List<string> nillableProperties = new List<string>();

            // Iterate through each row (skip header row if present)
            for (int row = firstRow + 1; row <= lastRow; row++)
            {
                // Column 0: Property Name (string)
                // Column 1: Optional flag (bool or string "true"/"false")
                string propertyName = cells[row, 0].StringValue?.Trim();
                if (string.IsNullOrEmpty(propertyName))
                    continue; // Skip rows without a name

                // Read the optional flag; handle both boolean and textual representations
                object optionalCell = cells[row, 1].Value;
                bool isOptional = false;

                if (optionalCell is bool boolVal)
                {
                    isOptional = boolVal;
                }
                else if (optionalCell != null)
                {
                    // Try to parse textual representation
                    bool.TryParse(optionalCell.ToString().Trim(), out isOptional);
                }

                // If the property is optional, consider it nillable
                if (isOptional)
                {
                    nillableProperties.Add(propertyName);
                }
            }

            // Generate the summary string
            string summary = nillableProperties.Count > 0
                ? "Nillable properties: " + string.Join(", ", nillableProperties)
                : "No nillable properties found.";

            // Output the summary
            Console.WriteLine(summary);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
