// Title: Rename the active worksheet to QuarterlyReport while preserving all cell data using Aspose.Cells for .NET
// AI Prompts: Assign "QuarterlyReport" to the Name property of the workbook's active Worksheet. | Refresh the workbook's ActiveSheetName property to reflect the new sheet title. | Persist the changes by saving the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# change name of current worksheet without affecting data | how to keep cell values when renaming a sheet in Aspose.Cells | set ActiveSheetName after renaming worksheet Aspose.Cells .NET | save workbook after worksheet rename using Aspose.Cells
// Tags: active worksheet name assignment Aspose.Cells | ActiveSheetName synchronization Aspose.Cells | saving workbook as xlsx after sheet rename Aspose.Cells | worksheet rename example C# Aspose.Cells

using Aspose.Cells;
using System;

namespace AsposeCellsExamples
{
    // Shows how to retrieve the active worksheet in an Aspose.Cells workbook, rename it to "QuarterlyReport", synchronize the ActiveSheetName property, and save the workbook as an .xlsx file while retaining all existing cell data.
    class RenameActiveWorksheet
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and worksheet renamed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create

            // Get the currently active worksheet
            Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Rename the active worksheet to "QuarterlyReport"
            activeSheet.Name = "QuarterlyReport";

            // Keep the ActiveSheetName property in sync (optional but recommended)
            workbook.Worksheets.ActiveSheetName = "QuarterlyReport";

            // Save the workbook; all existing cell data remains unchanged
            workbook.Save("RenamedWorkbook.xlsx", SaveFormat.Xlsx); // lifecycle: save
        }
    }
}
