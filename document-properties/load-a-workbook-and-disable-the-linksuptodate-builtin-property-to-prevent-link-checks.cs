// Title: Disable LinksUpToDate Property in an Excel Workbook with Aspose.Cells for .NET
// Description: Load an existing .xlsx file using Aspose.Cells, set the BuiltInDocumentProperties.LinksUpToDate flag to false to stop external link verification, and save the workbook.
// Keywords: Aspose.Cells LinksUpToDate | disable built‑in document property | prevent link checks Excel | C# Aspose.Cells workbook properties | turn off external link verification
// Common Searches: Aspose.Cells disable LinksUpToDate | set LinksUpToDate false C# | turn off link checking in Excel with Aspose | how to stop external link updates Aspose.Cells | disable built‑in document property LinksUpToDate
// Developer Intent: Turn off the LinksUpToDate flag so the workbook is saved without performing link validation.
// Use Cases: Distribute a workbook without triggering external data refreshes. | Create automated reports where link status is irrelevant. | Batch‑process files to remove link checks before archiving.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, disables LinksUpToDate, and saves it. | Show how to iterate over a folder of .xlsx files and set LinksUpToDate to false using Aspose.Cells for .NET. | Explain the impact of the LinksUpToDate property and how to turn it off programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Load an existing .xlsx file using Aspose.Cells, set the BuiltInDocumentProperties.LinksUpToDate flag to false to stop external link verification, and save the workbook.
    class DisableLinksUpToDate
    {
        static void Main()
        {
            // Load an existing workbook from a file
            // (uses the Workbook(string) constructor – the standard create/load rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Disable the LinksUpToDate built‑in property to prevent link checks
            // (access the BuiltInDocumentProperties collection and set the property)
            workbook.BuiltInDocumentProperties.LinksUpToDate = false;

            // Save the modified workbook to a new file
            // (uses the Workbook.Save(string) method – the standard save rule)
            workbook.Save("output.xlsx");

            // Optional: inform the user
            Console.WriteLine("LinksUpToDate property disabled and workbook saved as output.xlsx.");
        }
    }
}
