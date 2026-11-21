// Title: Open a password‑protected Excel workbook and remove its password with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load an encrypted XLSX file | how to clear LoadOptions.Password after opening a workbook in Aspose.Cells | remove workbook‑level password before saving with Aspose.Cells | save Excel file without password using Aspose.Cells .NET | unprotect Excel workbook programmatically with Aspose.Cells
// Tags: load encrypted XLSX using Aspose.Cells LoadOptions | clear LoadOptions password C# | strip workbook password Aspose.Cells | save unprotected Excel workbook Aspose.Cells | Aspose.Cells password removal workflow

using Aspose.Cells;
using System;

// The example loads a password‑protected XLSX workbook using LoadOptions, clears the password from LoadOptions and the workbook's Settings, and then saves the file as an unprotected Excel workbook.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";

        // Path for the workbook after removing the password
        string outputPath = "unprotected.xlsx";

        // Create LoadOptions and set the password required to open the file
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = "myPassword";

        // Load the workbook using the provided password
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Clear the password from LoadOptions after loading (no longer needed)
        loadOptions.Password = null;

        // Remove any workbook‑level password protection before saving
        workbook.Settings.Password = string.Empty;

        // Save the workbook without a password
        workbook.Save(outputPath);
    }
}
