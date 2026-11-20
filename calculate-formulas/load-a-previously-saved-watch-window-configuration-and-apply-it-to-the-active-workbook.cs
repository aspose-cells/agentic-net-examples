// Title: Load a saved Watch Window configuration and apply it to an active workbook using Aspose.Cells for .NET
// Common Searches: c# load watch window configuration with Aspose.Cells | apply saved watch window settings to an Excel workbook using Aspose.Cells .NET | how to handle missing WatchWindow class in Aspose.Cells version | load watch window XML file into a workbook with Aspose.Cells | save workbook after attempting to load watch window configuration in C#
// Tags: Aspose.Cells watch window loading | C# workbook save Aspose.Cells | missing WatchWindow API handling | watch window XML configuration Aspose.Cells | apply watch window settings .NET

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of a workbook and an optional WatchWindow configuration XML file, notes that the WatchWindow class is unavailable in the current Aspose.Cells API, provides a placeholder for loading the configuration when supported, and saves the workbook to a new file while handling errors.
class Program
{
    static void Main()
    {
        try
        {
            const string workbookPath = "MyWorkbook.xlsx";
            const string watchConfigPath = "WatchWindowConfig.xml";
            const string outputWorkbookPath = "MyWorkbook_WithWatchWindow.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // -----------------------------------------------------------------
            // NOTE: The WatchWindow class is not part of the current Aspose.Cells
            // API version. If you need to work with watch windows, refer to the
            // official documentation for the appropriate classes/methods.
            // The following block is kept for illustration and will be skipped
            // when the class is unavailable.
            // -----------------------------------------------------------------
            if (File.Exists(watchConfigPath))
            {
                Console.WriteLine($"Info: Watch window configuration file \"{watchConfigPath}\" found, but WatchWindow API is unavailable in this version.");
                // Placeholder for loading watch window configuration if supported:
                // watchWindow.Load(watchConfigPath);
            }
            else
            {
                Console.WriteLine($"Warning: Watch window configuration file \"{watchConfigPath}\" not found. Continuing without loading.");
            }

            // Save the workbook (persist any changes)
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved as \"{outputWorkbookPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
