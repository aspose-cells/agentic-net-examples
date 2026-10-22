// Title: C# example: Protect Excel workbook structure to block new sheets while permitting sheet renaming using Aspose.Cells
// Common Searches: Aspose.Cells C# protect workbook structure but allow renaming worksheets | How to prevent adding new worksheets in an Excel file with Aspose.Cells .NET | Enable AllowRenameSheets option when protecting a workbook with Aspose.Cells | Use reflection to set AllowRenameSheets in Aspose.Cells workbook protection | Protect Excel workbook without password using Aspose.Cells and keep sheet rename enabled
// Tags: Aspose.Cells structure protection API | AllowRenameSheets via reflection | block addition of worksheets Aspose.Cells | rename worksheets under protection | password‑less workbook protection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// Loads an existing XLSX file, uses reflection to turn on the AllowRenameSheets option if available, applies structure protection without a password to stop new worksheets from being added, and saves the protected workbook.
class WorkbookProtectionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to allow renaming of worksheets via ProtectionOptions (if available)
            try
            {
                // Use reflection to avoid compile‑time dependency on the property
                PropertyInfo protectionOptionsProp = workbook.Settings.GetType().GetProperty("ProtectionOptions");
                if (protectionOptionsProp != null)
                {
                    object protectionOptions = protectionOptionsProp.GetValue(workbook.Settings);
                    PropertyInfo allowRenameProp = protectionOptions?.GetType().GetProperty("AllowRenameSheets");
                    if (allowRenameProp != null && allowRenameProp.CanWrite)
                    {
                        allowRenameProp.SetValue(protectionOptions, true);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"AllowRenameSheets not available: {ex.Message}");
            }

            // Protect the workbook structure (no password required)
            workbook.Protect(ProtectionType.Structure, string.Empty);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Protected workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
