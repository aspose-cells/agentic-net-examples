// Title: Copy an XLSM workbook while preserving VBA UserForms using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy macro-enabled XLSM workbook without losing VBA UserForms | How to save an Excel file with macros using SaveFormat.Xlsm in .NET | Preserve VBA project when duplicating an XLSM file with Aspose.Cells | Copy Excel workbook and keep UserForm storage using Aspose.Cells API | C# code to duplicate a macro-enabled workbook while retaining macros
// Tags: copy XLSM workbook preserving VBA project Aspose.Cells | save workbook as Xlsm with macros C# | retain UserForm storage during workbook duplication | Aspose.Cells SaveFormat.Xlsm macro preservation | duplicate macro-enabled Excel file using .NET API

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro-enabled XLSM workbook with Aspose.Cells, which automatically includes the VBA project, and then saves it to a new location using SaveFormat.Xlsm, ensuring that all macros and UserForm storage remain intact.
class PreserveUserFormStorage
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsm";
            const string destinationPath = "destination.xlsm";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook (VBA project is loaded automatically)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Save the workbook to the destination path.
            // Saving as Xlsm automatically preserves the VBA project (including UserForms).
            sourceWorkbook.Save(destinationPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to \"{destinationPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
