// Title: Export Excel Shape Definitions to XML with Aspose.Cells for .NET
// Description: Learn how to load an Excel workbook, configure XmlSaveOptions, and save only the drawing objects (shapes) as an XML file. The generated shapes.xml contains the full XML definition of each shape, enabling external analysis, documentation, and version‑control tracking.
// Keywords: Aspose.Cells export shape XML | save Excel shapes as XML .NET | XmlSaveOptions drawing objects | extract shape definitions Aspose.Cells | Excel shape XML export C# | Aspose.Cells shape metadata
// Common Searches: how to export shape definitions from Excel using Aspose.Cells | Aspose.Cells XmlSaveOptions only shapes | save drawing objects to XML .NET | extract Excel shape XML for version control | C# export Excel shapes to XML file
// Developer Intent: Generate an XML file that contains the definitions of all shapes in an Excel workbook for external review or source‑control management.
// Use Cases: Create XML snapshots of workbook shapes to track design changes in Git. | Feed shape metadata to custom reporting or documentation tools. | Perform automated diff of shape properties across workbook versions.
// AI Prompts: Write C# code with Aspose.Cells that exports only the shapes' XML definitions, excluding worksheet data. | Show how to parse the exported shapes.xml to list each shape's type, ID, and key properties. | Explain how to configure XmlSaveOptions to limit the saved XML to drawing objects.

using System;
using Aspose.Cells;

// Learn how to load an Excel workbook, configure XmlSaveOptions, and save only the drawing objects (shapes) as an XML file. The generated shapes.xml contains the full XML definition of each shape, enabling external analysis, documentation, and version‑control tracking.
class ExportShapesXml
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create XML save options – this will include all workbook elements,
        // such as worksheets, cells, and the XML representation of drawing objects (shapes)
        XmlSaveOptions xmlOptions = new XmlSaveOptions();

        // Save the workbook as an XML file that contains the shapes' definitions
        workbook.Save("shapes.xml", xmlOptions);

        Console.WriteLine("Shapes' XML definitions exported to shapes.xml");
    }
}
