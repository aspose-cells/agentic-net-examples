using System;
using System.IO;
using Aspose.Cells;

class CheckDialogSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a dialog sheet (Add returns the index of the new sheet)
            int dialogSheetIndex = workbook.Worksheets.Add(SheetType.Dialog);
            Worksheet dialogWorksheet = workbook.Worksheets[dialogSheetIndex];
            dialogWorksheet.Name = "MyDialogSheet";

            // Determine whether the worksheet is a dialog sheet by checking its type
            bool isDialogSheet = dialogWorksheet.Type == SheetType.Dialog;

            // Output the result
            Console.WriteLine($"Worksheet \"{dialogWorksheet.Name}\" IsDialogSheet: {isDialogSheet}");

            // Save the workbook
            string outputPath = "DialogSheetDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}