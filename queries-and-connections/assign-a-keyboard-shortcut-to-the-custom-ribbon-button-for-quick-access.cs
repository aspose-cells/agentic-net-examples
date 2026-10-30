// Title: How to add a "ShortcutSheet" worksheet to an existing XLSX workbook only if it doesn't exist using Aspose.Cells for .NET
// Common Searches: asp.net add worksheet to existing xlsx only if missing using Aspose.Cells | c# check if a worksheet exists before adding it with Aspose.Cells | load an existing workbook and overwrite the same file with Aspose.Cells C# | validate workbook file path before loading it in Aspose.Cells | create a new workbook automatically when the file is not found using Aspose.Cells
// Tags: add worksheet if missing Aspose.Cells | load existing workbook C# Aspose.Cells | save workbook in place Aspose.Cells | validate workbook path C# | create workbook when file not found Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example shows how to use Aspose.Cells for .NET to load an existing XLSX file, add a worksheet named "ShortcutSheet" only if it is absent, and save the changes back to the same file while handling missing files and invalid paths.
    public class RibbonShortcutHandler
    {
        // Constructor (no Excel interop dependency)
        public RibbonShortcutHandler()
        {
        }

        // Adds a worksheet named "ShortcutSheet" to the specified workbook file.
        public void AddShortcutSheet(string workbookPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(workbookPath))
                    throw new ArgumentException("Workbook path must be provided.", nameof(workbookPath));

                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("The specified workbook file does not exist.", workbookPath);

                // Load the workbook using Aspose.Cells
                var workbook = new Workbook(workbookPath);

                // Add a new worksheet named "ShortcutSheet" if it doesn't already exist
                if (workbook.Worksheets["ShortcutSheet"] == null)
                {
                    workbook.Worksheets.Add("ShortcutSheet");
                }

                // Save changes back to the original file
                workbook.Save(workbookPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Log the error
                Console.Error.WriteLine($"Error in AddShortcutSheet: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the workbook (adjust as needed)
                string workbookPath = "Sample.xlsx";

                // Ensure the workbook exists; create a new one if it doesn't
                if (!File.Exists(workbookPath))
                {
                    var wb = new Workbook();
                    wb.Save(workbookPath, SaveFormat.Xlsx);
                }

                var handler = new RibbonShortcutHandler();
                handler.AddShortcutSheet(workbookPath);

                Console.WriteLine("Shortcut sheet added successfully.");
            }
            catch (Exception ex)
            {
                // Log any unhandled exceptions
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
