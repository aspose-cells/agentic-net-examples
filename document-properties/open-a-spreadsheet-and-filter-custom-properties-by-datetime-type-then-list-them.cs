// Title: C# – List DateTime Custom Document Properties in an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells for .NET, accesses the CustomDocumentPropertyCollection, filters properties whose value is a DateTime, and prints each name with its ISO‑8601 timestamp. The workbook can be saved afterwards if needed.
// Keywords: Aspose.Cells | .NET | C# | custom document properties | DateTime filter | Excel metadata | list properties | ISO 8601
// Common Searches: Aspose.Cells read DateTime custom properties | C# filter Excel custom document properties by type | list DateTime metadata in Excel using Aspose | how to get custom property values as ISO 8601 with Aspose.Cells | extract Excel custom properties .NET
// Developer Intent: Extract and display only the custom document properties that contain DateTime values from an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create an audit log of all timestamp‑based custom properties in a workbook. | Synchronize Excel file timestamps with external databases or APIs. | Validate presence of required DateTime metadata before running data‑processing pipelines.
// AI Prompts: Generate C# code with Aspose.Cells that reads an Excel file and prints custom document properties whose values are DateTime objects in ISO‑8601 format. | Write a method that returns a Dictionary<string, DateTime> of DateTime‑type custom properties from a workbook, ignoring non‑DateTime entries. | Explain how to extend the sample to filter DateTime properties within a specific date range using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an Excel file using Aspose.Cells for .NET, accesses the CustomDocumentPropertyCollection, filters properties whose value is a DateTime, and prints each name with its ISO‑8601 timestamp. The workbook can be saved afterwards if needed.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the custom document properties collection
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Iterate through all custom properties and output those of DateTime type
        foreach (DocumentProperty prop in customProps)
        {
            // Check if the property's value is a DateTime instance
            if (prop.Value is DateTime dtValue)
            {
                // Output the property name and its DateTime value in ISO 8601 format
                Console.WriteLine($"{prop.Name}: {dtValue:O}");
            }
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
