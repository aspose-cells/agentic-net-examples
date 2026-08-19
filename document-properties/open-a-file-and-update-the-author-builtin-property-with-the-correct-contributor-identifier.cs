// Title: Set the Author Built‑in Property of an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: A C# sample that loads an existing .xlsx file via Aspose.Cells, assigns a new Author value such as a contributor ID, and saves the workbook with the revised metadata.
// Keywords: Aspose.Cells C# author property | Excel built‑in document properties | modify workbook metadata | set contributor ID Excel | update Excel Author programmatically | Aspose.Cells document properties API
// Common Searches: how to change author in Excel using Aspose.Cells C# | Aspose.Cells set built‑in document properties .NET | update contributor identifier in workbook metadata | C# code to modify Excel Author field | batch update Excel author property Aspose
// Developer Intent: Open an existing Excel file and assign a specific contributor identifier to its Author built‑in property.
// Use Cases: Ensure consistent author attribution across automatically generated reports. | Embed contributor IDs for audit trails and regulatory compliance. | Run bulk updates on Excel files so the Author field reflects the processing system or user.
// AI Prompts: Write C# code with Aspose.Cells that reads an .xlsx file, sets the Author property from a variable, and saves the changes. | Show how to update several built‑in document properties (Author, Title, Subject) in one pass using Aspose.Cells for .NET. | Explain how to retrieve the current Author value, compare it, and then overwrite it with a new contributor ID in an Excel workbook.

using System;
using Aspose.Cells;

// A C# sample that loads an existing .xlsx file via Aspose.Cells, assigns a new Author value such as a contributor ID, and saves the workbook with the revised metadata.
class UpdateAuthorProperty
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Update the built‑in Author property with the contributor identifier
        workbook.BuiltInDocumentProperties.Author = "ContributorID123";

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}
