// Title: Apply a Project Prefix to Custom XML Part IDs in Aspose.Cells for .NET
// Description: Demonstrates how to add a custom XML part to an Aspose.Cells workbook, assign a GUID as its ID, prepend a project-specific prefix, save the file, and confirm that the prefixed ID persists after reloading.
// Keywords: Aspose.Cells | .NET | CustomXmlPart | ID prefix | project identifier | Excel workbook | GUID | XML part naming | add custom XML part | modify custom XML part ID
// Common Searches: Aspose.Cells add custom XML part with prefixed ID | C# set custom XML part ID prefix in Excel workbook | How to prepend project code to CustomXmlPart ID using Aspose.Cells | Persist custom XML part ID after saving workbook | Rename CustomXmlPart ID in Aspose.Cells .NET
// Developer Intent: Add a custom XML part to a workbook and prepend a project‑specific string to its identifier.
// Use Cases: Enforce a uniform naming scheme for all custom XML parts across multiple workbooks. | Embed project metadata directly in the XML part ID for downstream automation. | Validate that the prefixed identifier remains unchanged after the workbook is saved and reopened.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom XML part, sets its ID using a given project prefix, saves the workbook, and prints the ID after loading. | Show how to iterate through every CustomXmlPart in a workbook and apply the same project prefix to each part's ID. | Explain the steps to retrieve, modify, and persist the ID of an existing CustomXmlPart in an Aspose.Cells workbook.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates how to add a custom XML part to an Aspose.Cells workbook, assign a GUID as its ID, prepend a project-specific prefix, save the file, and confirm that the prefixed ID persists after reloading.
class ApplyCustomXmlPartNamingConvention
{
    static void Main()
    {
        // Project identifier to prefix custom XML part IDs
        const string projectPrefix = "ProjA_";

        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Sample XML data for the custom part
        string xmlData = "<root><item>Sample</item></root>";
        byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);

        // Add the custom XML part to the workbook (add method rule)
        int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);
        CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];

        // Assign an initial GUID as the ID
        customPart.ID = Guid.NewGuid().ToString();

        // Apply naming convention: prefix the ID with the project identifier
        customPart.ID = projectPrefix + customPart.ID;

        // Save the workbook (save rule)
        string outputPath = "CustomXmlPartPrefixed.xlsx";
        workbook.Save(outputPath);

        // Load the workbook to verify the prefixed ID (load rule)
        Workbook loadedWorkbook = new Workbook(outputPath);
        CustomXmlPart loadedPart = loadedWorkbook.CustomXmlParts[0];
        Console.WriteLine("Prefixed Custom XML Part ID: " + loadedPart.ID);
    }
}
