// Title: How to assign a custom integer TabId to a worksheet and save the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet TabId to a specific integer | change Excel sheet tab identifier using Aspose.Cells .NET | save workbook after modifying worksheet TabId programmatically | load workbook, assign TabId, and verify with Aspose.Cells | assign TabId 107 to a target worksheet in Aspose.Cells example
// Tags: worksheet TabId assignment Aspose.Cells | Aspose.Cells set worksheet tab identifier | save workbook after TabId change .NET | load or create workbook Aspose.Cells C# | verify worksheet TabId persistence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AssignWorksheetTabId
{
    // The example loads an existing workbook (or creates a new one), selects a worksheet named "TargetSheet" or the first sheet, assigns the integer value 107 to its TabId property, saves the workbook, and then reloads it to confirm that the TabId was persisted.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths.
            // If an input file is provided and exists, it will be loaded; otherwise a new workbook is created.
            string inputPath = args.Length > 0 ? args[0] : string.Empty;
            string outputPath = args.Length > 1 ? args[1] : "WorkbookWithTabId.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if the file exists; otherwise create a new one.
                if (!string.IsNullOrWhiteSpace(inputPath) && File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine("No valid input file provided. Creating a new workbook.");
                    workbook = new Workbook(); // Default format is Xlsx
                }

                // Choose the worksheet to modify.
                // If a worksheet named "TargetSheet" exists, use it; otherwise use the first worksheet.
                const string targetSheetName = "TargetSheet";
                Worksheet worksheet = workbook.Worksheets[targetSheetName];

                if (worksheet != null)
                {
                    Console.WriteLine($"Found existing worksheet \"{targetSheetName}\".");
                }
                else
                {
                    worksheet = workbook.Worksheets[0];
                    Console.WriteLine($"Using first worksheet \"{worksheet.Name}\" (index 0).");
                }

                // Assign a new TabId. The value must be a positive integer.
                int newTabId = 107;
                worksheet.TabId = newTabId;
                Console.WriteLine($"Assigned TabId = {newTabId} to worksheet \"{worksheet.Name}\".");

                // Save the workbook.
                Console.WriteLine($"Saving workbook to \"{outputPath}\"...");
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully.");

                // Verify by reloading and reading the TabId.
                Workbook verificationWorkbook = new Workbook(outputPath);
                Worksheet verificationSheet = verificationWorkbook.Worksheets[worksheet.Index];
                Console.WriteLine($"Verification: Worksheet \"{verificationSheet.Name}\" has TabId = {verificationSheet.TabId}.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
