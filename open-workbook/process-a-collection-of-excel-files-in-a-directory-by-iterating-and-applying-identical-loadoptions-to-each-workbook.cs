// Title: Load multiple Excel workbooks from a folder with identical LoadOptions using Aspose.Cells for .NET
// Common Searches: C# load all .xlsx files in a directory with the same LoadOptions using Aspose.Cells | How to batch process Excel workbooks with password protection in Aspose.Cells .NET | Iterate through a folder and apply LoadOptions to each workbook in Aspose.Cells C# | Aspose.Cells load multiple workbooks with custom LoadOptions and save changes | Read and save every Excel file in a folder using Aspose.Cells LoadOptions
// Tags: batch loading Excel workbooks with Aspose.Cells LoadOptions | apply uniform LoadOptions to multiple .xlsx files | enumerate Excel files in folder C# Aspose.Cells | filter .xls .xlsx .xlsm extensions Aspose.Cells | save modified workbook to original path Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all files in a specified directory, skips non‑Excel extensions, loads each workbook with a shared LoadOptions instance, allows custom processing, and saves the workbook back to its original location.
class Program
{
    static void Main()
    {
        // Path to the directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Define the LoadOptions that will be applied to every workbook
        LoadOptions loadOptions = new LoadOptions
        {
            // Example options (customize as needed):
            // LoadFormat = LoadFormat.Xlsx, // force a specific format if required
            // Password = "yourPassword",   // use if files are password‑protected
        };

        // Retrieve all files in the directory
        string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only Excel files (xls, xlsx, xlsm)
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                continue;

            // Load the workbook using the predefined LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // -------------------------------------------------
            // Insert any workbook processing logic here.
            // For example: workbook.CalculateFormula();
            // -------------------------------------------------

            // Save the workbook back to the original file (or specify a new path)
            workbook.Save(filePath);
        }
    }
}
