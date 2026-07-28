// Title: List DateTime Custom Document Properties in an Excel Workbook with Aspose.Cells for .NET
// Description: This C# example shows how to open an Excel file using Aspose.Cells, access its CustomDocumentPropertyCollection, filter for properties of type DateTime, and output each property's name together with an ISO‑8601 timestamp. The workbook can be saved afterward if modifications are required.
// Keywords: Aspose.Cells | C# | Excel | custom document properties | DateTime | filter properties | list metadata | property type DateTime | Workbook.CustomDocumentProperties | retrieve Excel metadata
// Common Searches: Aspose.Cells list DateTime custom properties | C# filter Excel custom document properties by type | How to read DateTime metadata from an Excel file using Aspose.Cells | Get custom document properties of type DateTime in .NET | Extract timestamp properties from workbook with Aspose.Cells
// Developer Intent: The developer wants to open an Excel workbook and display only the custom document properties that are of DateTime type.
// Use Cases: Generate an audit trail by extracting timestamp metadata stored as custom properties in financial spreadsheets. | Validate the presence of required DateTime metadata (e.g., creation or review dates) before processing workbook data. | Synchronize version dates across multiple Excel files by reading their DateTime custom properties and updating a central repository.
// AI Prompts: Provide C# code using Aspose.Cells to extract all custom document properties of type DateTime and return them as a dictionary of name/value pairs. | Show how to modify the DateTime value of a specific custom property and save the workbook with Aspose.Cells. | Explain how to safely handle missing or incorrectly typed custom properties when enumerating them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This C# example shows how to open an Excel file using Aspose.Cells, access its CustomDocumentPropertyCollection, filter for properties of type DateTime, and output each property's name together with an ISO‑8601 timestamp. The workbook can be saved afterward if modifications are required.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        Console.WriteLine("DateTime custom properties:");

        // Iterate through all custom properties
        foreach (DocumentProperty prop in customProps)
        {
            // Filter only those whose type is DateTime
            if (prop.Type == PropertyType.DateTime)
            {
                // Retrieve the DateTime value and display it
                DateTime dt = prop.ToDateTime();
                Console.WriteLine($"{prop.Name}: {dt:O}");
            }
        }

        // Save the workbook (optional, to keep any modifications)
        workbook.Save("output.xlsx");
    }
}
