// Title: Copy VBA UserForms from a macro‑enabled Excel template into a new workbook using Aspose.Cells for .NET
// Common Searches: aspnet copy userforms from xlsm template to new workbook using Aspose.Cells | c# duplicate VBA project when cloning a macro-enabled Excel file with Aspose.Cells | how to preserve userforms while saving a cloned workbook as xlsm in Aspose.Cells | Aspose.Cells copy method retain VBA forms in new macro-enabled workbook C#
// Tags: Aspose.Cells Workbook.Copy macro-enabled | preserve VBA UserForms Xlsm | clone Excel workbook with VBA C# | copy VBA project Aspose.Cells | save macro-enabled workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled Excel template (XLSM) that contains VBA UserForms, creates a new workbook, uses Workbook.Copy to duplicate the source workbook—including its VBA project and UserForms—and saves the result as a new XLSM file, with error handling for missing files.
class Program
{
    static void Main()
    {
        const string sourcePath = "SourceTemplate.xlsm";
        const string targetPath = "TargetWithUserForms.xlsm";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook (macro‑enabled) which already contains the UserForms
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook by cloning the source workbook.
            // This preserves the VBA project (including UserForms) in the new file.
            Workbook targetWorkbook = new Workbook();
            targetWorkbook.Copy(sourceWorkbook);

            // Save the target workbook as a macro‑enabled file.
            targetWorkbook.Save(targetPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{targetPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
