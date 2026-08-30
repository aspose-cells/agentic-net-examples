// Title: Copy selected worksheets from a source Excel workbook to a new workbook using Aspose.Cells Worksheet.Copy in C#
// AI Prompts: Copy the worksheets "Sheet1" and "Data" from Source.xlsx into a newly created workbook while preserving all formatting using Aspose.Cells in C#. | Update the code to accept a runtime list of worksheet names and copy each existing sheet from the source workbook to the target workbook with Worksheet.Copy. | Add robust handling for missing worksheets and ensure the merged workbook is saved as Target.xlsx in the correct Excel format.
// Common Searches: Aspose.Cells C# copy specific sheets from one workbook to another preserving formatting | How to use Worksheet.Copy to duplicate selected worksheets into a new Excel file in C# | Programmatically merge multiple worksheets into a new workbook with Aspose.Cells | C# example for copying sheets by name from source Excel to target workbook using Aspose.Cells
// Tags: Worksheet.Copy selected sheets Aspose.Cells | copy worksheets between workbooks C# | preserve formatting when copying Excel sheets Aspose | dynamic worksheet list copying Aspose.Cells | error handling missing worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads Source.xlsx, creates an empty workbook, copies the specified worksheets (e.g., "Sheet1" and "Data") using Worksheet.Copy to retain content and formatting, handles absent sheets gracefully, and saves the result as Target.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string targetPath = "Target.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty target workbook
            Workbook targetWorkbook = new Workbook();
            targetWorkbook.Worksheets.Clear();

            // Worksheets to copy
            string[] sheetsToCopy = { "Sheet1", "Data" };

            foreach (string sheetName in sheetsToCopy)
            {
                // Get source worksheet; skip if not found
                Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
                if (sourceSheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found in source workbook.");
                    continue;
                }

                // Add a new worksheet with the same name to the target workbook
                Worksheet destSheet = targetWorkbook.Worksheets.Add(sheetName);

                // Copy contents and formatting
                destSheet.Copy(sourceSheet);
            }

            // Save the target workbook
            targetWorkbook.Save(targetPath);
            Console.WriteLine($"Target workbook saved to {targetPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
