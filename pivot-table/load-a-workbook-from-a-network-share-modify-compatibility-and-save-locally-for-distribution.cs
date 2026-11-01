// Title: Load an Excel workbook from a UNC network share and save it locally with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load Excel workbook from UNC network share | How to save a workbook to a local folder after reading it from a network share using Aspose.Cells | Verify file existence on a Windows network share before opening with Aspose.Cells | Create destination directory automatically before saving workbook in C# | Apply optional compatibility settings to an Aspose.Cells workbook before export
// Tags: UNC path workbook loading Aspose.Cells | local save operation Aspose.Cells | network share file verification C# | destination directory creation Aspose.Cells | optional compatibility settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for an Excel file on a UNC network share, loads it with Aspose.Cells' Workbook class, optionally adjusts compatibility options, ensures the target local directory exists, and saves the workbook as an .xlsx file to a specified local path, handling any errors that occur.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook on a network share
                string networkPath = @"\\Server\Share\Folder\SourceWorkbook.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(networkPath))
                {
                    Console.WriteLine($"Source workbook not found: {networkPath}");
                    return;
                }

                // Load the workbook from the network location
                Workbook workbook = new Workbook(networkPath);

                // NOTE: Compatibility options are not required for this example.
                // If needed, they can be set using workbook.Settings.CompatibilityOptions
                // in versions of Aspose.Cells that support this property.

                // Path to save the modified workbook locally for distribution
                string localPath = @"C:\Distribution\ModifiedWorkbook.xlsx";

                // Ensure the target directory exists
                string? localDir = Path.GetDirectoryName(localPath);
                if (!string.IsNullOrEmpty(localDir) && !Directory.Exists(localDir))
                {
                    Directory.CreateDirectory(localDir);
                }

                // Save the workbook to the local path
                workbook.Save(localPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {localPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
