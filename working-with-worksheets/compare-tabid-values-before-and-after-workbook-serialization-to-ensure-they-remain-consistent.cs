// Title: Check Worksheet TabId Persistence After Save and Load with Aspose.Cells for .NET
// Description: Demonstrates how to set a worksheet's TabId, serialize the workbook to XLSX, reload it, and verify that the TabId value remains unchanged, with proper cleanup of resources.
// Keywords: Aspose.Cells TabId verification | worksheet TabId after save | preserve TabId XLSX | TabId consistency Aspose | C# Aspose.Cells workbook serialization
// Common Searches: Aspose.Cells keep worksheet TabId after saving | verify TabId value after workbook reload .NET | how to test TabId persistence in Excel file | compare TabId before and after Aspose.Cells save
// Developer Intent: Confirm that the TabId property of a worksheet is retained unchanged when a workbook is saved and later reloaded using Aspose.Cells for .NET.
// Use Cases: Automated CI test to ensure custom sheet identifiers survive serialization. | Migration script that validates sheet IDs are not altered during format conversion. | Debug routine for downstream processes that rely on stable TabId values.
// AI Prompts: Write C# code with Aspose.Cells that assigns a TabId, saves the workbook, reloads it, and asserts the TabId is unchanged. | Explain the internal storage of the TabId property in an XLSX file and which save options might affect it. | Create an MSTest/NUnit unit test that checks TabId consistency after workbook serialization using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTabIdComparison
{
    // Demonstrates how to set a worksheet's TabId, serialize the workbook to XLSX, reload it, and verify that the TabId value remains unchanged, with proper cleanup of resources.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a known TabId value
            int originalTabId = 12345;
            worksheet.TabId = originalTabId;

            // Define a temporary file path for serialization
            string tempFilePath = Path.Combine(Path.GetTempPath(), "TabIdTest.xlsx");

            // Save the workbook (using the provided save rule)
            workbook.Save(tempFilePath, SaveFormat.Xlsx);

            // Load the saved workbook (using the provided load rule)
            Workbook loadedWorkbook = new Workbook(tempFilePath);

            // Retrieve the TabId from the loaded worksheet
            int loadedTabId = loadedWorkbook.Worksheets[0].TabId;

            // Compare the original and loaded TabId values
            if (originalTabId == loadedTabId)
            {
                Console.WriteLine($"Success: TabId is consistent ({originalTabId}).");
            }
            else
            {
                Console.WriteLine($"Failure: Original TabId ({originalTabId}) != Loaded TabId ({loadedTabId}).");
            }

            // Clean up temporary file
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

            // Dispose workbooks
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
