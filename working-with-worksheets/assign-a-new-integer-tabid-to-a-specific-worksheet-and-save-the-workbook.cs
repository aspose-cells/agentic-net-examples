// Title: Set and Persist a Worksheet TabId with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to assign a custom integer TabId to a worksheet, save the workbook, reload it, and verify that the TabId value is retained using Aspose.Cells for .NET.
// Keywords: Aspose.Cells TabId | Worksheet TabId C# | set worksheet TabId .NET | save workbook TabId | load workbook TabId | Aspose.Cells worksheet identifier | C# Excel TabId example | persist worksheet ID | Aspose.Cells API TabId property | Excel custom worksheet ID
// Common Searches: how to set worksheet TabId Aspose.Cells | persist TabId after saving workbook C# | read TabId from loaded Excel file | Aspose.Cells assign unique worksheet identifier | C# example for Worksheet.TabId
// Developer Intent: The developer needs to assign a numeric TabId to a specific worksheet and ensure the value is stored in the file so it can be retrieved later.
// Use Cases: Create a stable numeric key for each sheet to synchronize with external databases. | Track worksheet versions across deployments without relying on sheet names. | Map worksheets to API resources or business objects using a consistent identifier.
// AI Prompts: Show C# code that sets Worksheet.TabId, saves the workbook, and reads the value back. | Explain when to use TabId instead of Worksheet.Name in Aspose.Cells. | Provide a step‑by‑step guide to verify that a custom TabId persists after reopening the file.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Demonstrates how to assign a custom integer TabId to a worksheet, save the workbook, reload it, and verify that the TabId value is retained using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign a new TabId value
            worksheet.TabId = 12345;

            // Define output file path
            string outputPath = "TabIdDemo.xlsx";

            // Save the workbook with the updated TabId
            workbook.Save(outputPath);

            // Load the saved workbook to verify the TabId persists
            Workbook loadedWorkbook = new Workbook(outputPath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Loaded Worksheet TabId: " + loadedWorksheet.TabId);
        }
    }
}
