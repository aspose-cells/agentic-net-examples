using System;
using Aspose.Cells;

namespace AsposeCellsTabIdVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a known TabId value
            int originalTabId = 12345;
            sheet.TabId = originalTabId;

            // Define a temporary file path for serialization
            string tempFile = "TabIdTest.xlsx";

            // Save the workbook to disk
            workbook.Save(tempFile);
            // Dispose the original workbook to release resources
            workbook.Dispose();

            // Load the workbook back from the saved file
            Workbook loadedWorkbook = new Workbook(tempFile);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Retrieve the TabId after deserialization
            int loadedTabId = loadedSheet.TabId;

            // Compare the original and loaded TabId values
            if (originalTabId == loadedTabId)
            {
                Console.WriteLine($"Success: TabId is consistent ({originalTabId}).");
            }
            else
            {
                Console.WriteLine($"Failure: Original TabId ({originalTabId}) != Loaded TabId ({loadedTabId}).");
            }

            // Clean up
            loadedWorkbook.Dispose();
        }
    }
}