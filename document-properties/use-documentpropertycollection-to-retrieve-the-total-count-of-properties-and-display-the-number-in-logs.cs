// Title: Count Built‑In and Custom Document Properties with Aspose.Cells for .NET
// Description: Creates an empty workbook, reads the BuiltInDocumentPropertyCollection and CustomDocumentPropertyCollection, logs each collection's Count, and optionally saves the file.
// Keywords: Aspose.Cells document property count | built‑in properties Aspose.Cells .NET | custom document properties count | log property totals C# | Workbook property collection Aspose
// Common Searches: how to count built‑in document properties Aspose.Cells | total custom document properties in a workbook .NET | log document property numbers using Aspose.Cells | retrieve document property collection count C# | Aspose.Cells count properties example
// Developer Intent: Obtain the number of built‑in and custom document properties in a workbook and output the values.
// Use Cases: Verify default built‑in properties after creating a new workbook. | Audit a batch of workbooks to ensure required custom metadata is present. | Log property counts during automated report generation for compliance tracking.
// AI Prompts: Generate C# code with Aspose.Cells that prints the count of built‑in and custom document properties. | Show how to iterate over all custom document properties after counting them and display each name/value pair. | Explain the steps to add a new custom document property, then recalculate and log the updated property count.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDocumentPropertyCountDemo
{
    // Creates an empty workbook, reads the BuiltInDocumentPropertyCollection and CustomDocumentPropertyCollection, logs each collection's Count, and optionally saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook)
            Workbook workbook = new Workbook();

            // Retrieve the built‑in document properties collection
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            // Get the total number of built‑in properties
            int builtInCount = builtInProps.Count;

            // Log the count
            Console.WriteLine($"Total built‑in document properties: {builtInCount}");

            // Retrieve the custom document properties collection
            CustomDocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;

            // Get the total number of custom properties
            int customCount = customProps.Count;

            // Log the count
            Console.WriteLine($"Total custom document properties: {customCount}");

            // (Optional) Save the workbook to demonstrate lifecycle usage
            workbook.Save("DocumentPropertyCountDemo.xlsx");
        }
    }
}
