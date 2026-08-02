// Title: Verify Worksheet TabId Consistency After Save‑Load with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to assign a TabId to a worksheet, save the workbook to XLSX, reload it, and compare the TabId values to ensure they remain unchanged through serialization using Aspose.Cells for .NET.
// Keywords: Aspose.Cells TabId | worksheet TabId C# | Excel serialization consistency | save load TabId Aspose | Aspose.Cells .NET worksheet properties | TabId round‑trip test
// Common Searches: Aspose.Cells keep worksheet TabId after saving | C# compare TabId before and after workbook serialization | verify TabId persistence in Excel file Aspose | how to test worksheet TabId consistency with Aspose.Cells | TabId value changes after save load Aspose.Cells
// Developer Intent: Confirm that a worksheet's TabId does not change when the workbook is saved and reopened.
// Use Cases: Automated regression test to ensure custom TabId values survive round‑trip saves. | Validation step in a data‑processing pipeline that relies on stable worksheet identifiers. | Debugging scenario where unexpected TabId changes indicate a serialization issue.
// AI Prompts: Write C# code with Aspose.Cells that sets a worksheet's TabId, saves the workbook, reloads it, and asserts equality of the TabId values. | Create an xUnit test that verifies TabId consistency after workbook serialization using Aspose.Cells for .NET. | Explain the internal storage of TabId in an XLSX file and list factors that could alter it during save or load.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to assign a TabId to a worksheet, save the workbook to XLSX, reload it, and compare the TabId values to ensure they remain unchanged through serialization using Aspose.Cells for .NET.
    public class WorksheetTabIdConsistencyDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a specific TabId value
            int originalTabId = 12345;
            sheet.TabId = originalTabId;

            // Store the TabId before saving
            int beforeSaveTabId = sheet.TabId;
            Console.WriteLine("TabId before save: " + beforeSaveTabId);

            // Save the workbook to a temporary file
            string tempPath = Path.Combine(Path.GetTempPath(), "TabIdDemo.xlsx");
            workbook.Save(tempPath, SaveFormat.Xlsx);

            // Ensure the file exists before loading
            if (!File.Exists(tempPath))
            {
                Console.WriteLine("Error: Temporary file was not created.");
                return;
            }

            // Load the workbook from the saved file
            Workbook loadedWorkbook = new Workbook(tempPath);

            // Retrieve the same worksheet (index 0) from the loaded workbook
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            int afterLoadTabId = loadedSheet.TabId;
            Console.WriteLine("TabId after load: " + afterLoadTabId);

            // Compare the TabId values
            if (beforeSaveTabId == afterLoadTabId)
            {
                Console.WriteLine("Success: TabId values are consistent after serialization.");
            }
            else
            {
                Console.WriteLine("Failure: TabId values differ after serialization.");
            }

            // Clean up temporary file
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Warning: Unable to delete temporary file. " + ex.Message);
            }

            // Dispose workbooks
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
