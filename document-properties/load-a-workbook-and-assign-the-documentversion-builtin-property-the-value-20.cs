// Title: Set DocumentVersion Built‑In Property in an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Load an existing workbook, assign the built‑in DocumentVersion property the value "2.0" via Workbook.BuiltInDocumentProperties, and save the file as Xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells DocumentVersion | set built‑in document property C# | Excel metadata update .NET | Workbook BuiltInDocumentProperties | modify Excel file version | C# Aspose.Cells example
// Common Searches: Aspose.Cells set DocumentVersion C# | How to change built‑in document properties in Excel with Aspose | Update Excel file version property programmatically .NET | C# code to modify DocumentVersion property Aspose.Cells | Excel workbook metadata Aspose.Cells tutorial
// Developer Intent: Load an Excel workbook, set its DocumentVersion built‑in property to "2.0", and save the updated file.
// Use Cases: Maintain consistent version tags across generated reports for traceability. | Embed version information in workbooks to satisfy audit and compliance requirements. | Prepare Excel files with a specific version identifier before distribution to downstream systems.
// AI Prompts: Show how to update other built‑in properties such as Author, Title, and Subject with Aspose.Cells. | Create a script that batch‑processes multiple workbooks to set several document properties at once. | Demonstrate how to read back the DocumentVersion property after saving to confirm the change.

using System;
using Aspose.Cells;

// Load an existing workbook, assign the built‑in DocumentVersion property the value "2.0" via Workbook.BuiltInDocumentProperties, and save the file as Xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Assign the built‑in DocumentVersion property the value "2.0"
        workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
