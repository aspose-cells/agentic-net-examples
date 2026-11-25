// Title: Check that worksheet TabId values stay the same after saving and loading an XLSX file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify worksheet TabId after saving to XLSX | how to ensure Excel sheet TabId is preserved when using Aspose.Cells | compare TabId values before and after workbook serialization in .NET | C# Aspose.Cells load workbook and validate TabId consistency | detect TabId changes after saving Excel file with Aspose.Cells
// Tags: worksheet TabId persistence Aspose.Cells | Aspose.Cells save workbook preserve TabId | validate TabId consistency after XLSX serialization | C# retrieve worksheet TabId array | Aspose.Cells compare original and loaded TabId values

using System;
using System.IO;
using Aspose.Cells;

namespace TabIdConsistencyCheck
{
    // The example creates a workbook with three worksheets, assigns explicit TabId values (101, 202, 303), saves the workbook to a temporary XLSX file, reloads it, extracts the TabId values from both the original and loaded workbooks, and compares them to confirm that the TabId values remain identical after serialization.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for temporary workbook file
            string tempFilePath = Path.Combine(Path.GetTempPath(), "TabIdConsistencyCheck.xlsx");

            try
            {
                // -------------------------------------------------
                // Step 1: Create a workbook and set known TabId values
                // -------------------------------------------------
                Workbook originalWorkbook = new Workbook();

                // Ensure we have at least three worksheets for the test
                Worksheet sheet1 = originalWorkbook.Worksheets[0]; // default sheet
                sheet1.Name = "FirstSheet";
                sheet1.TabId = 101; // explicit TabId

                Worksheet sheet2 = originalWorkbook.Worksheets.Add("SecondSheet");
                sheet2.TabId = 202;

                Worksheet sheet3 = originalWorkbook.Worksheets.Add("ThirdSheet");
                sheet3.TabId = 303;

                // Store the original TabId values for later comparison
                int[] originalTabIds = GetTabIds(originalWorkbook);

                // -------------------------------------------------
                // Step 2: Serialize (save) the workbook to a file
                // -------------------------------------------------
                originalWorkbook.Save(tempFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {tempFilePath}");

                // -------------------------------------------------
                // Step 3: Load the workbook back from the file
                // -------------------------------------------------
                Workbook loadedWorkbook = new Workbook(tempFilePath);
                int[] loadedTabIds = GetTabIds(loadedWorkbook);

                // -------------------------------------------------
                // Step 4: Compare TabId values before and after serialization
                // -------------------------------------------------
                bool allMatch = true;
                for (int i = 0; i < originalTabIds.Length; i++)
                {
                    Console.WriteLine($"Worksheet {i} - Original TabId: {originalTabIds[i]}, Loaded TabId: {loadedTabIds[i]}");
                    if (originalTabIds[i] != loadedTabIds[i])
                    {
                        allMatch = false;
                        Console.WriteLine($"Mismatch detected on worksheet index {i}.");
                    }
                }

                if (allMatch)
                {
                    Console.WriteLine("SUCCESS: All TabId values are consistent after serialization.");
                }
                else
                {
                    Console.WriteLine("FAILURE: One or more TabId values changed after serialization.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the TabId consistency check:");
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // Clean up the temporary file if it exists
                try
                {
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                        Console.WriteLine($"Temporary file deleted: {tempFilePath}");
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"Failed to delete temporary file: {cleanupEx.Message}");
                }
            }
        }

        /// <param name="workbook">The workbook to inspect.</param>
        /// <returns>An array of TabId integers ordered by worksheet index.</returns>
        private static int[] GetTabIds(Workbook workbook)
        {
            int sheetCount = workbook.Worksheets.Count;
            int[] tabIds = new int[sheetCount];
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                tabIds[i] = ws.TabId;
            }
            return tabIds;
        }
    }
}
