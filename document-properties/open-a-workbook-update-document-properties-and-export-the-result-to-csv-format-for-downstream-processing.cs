// Title: Update Excel workbook document properties via reflection and save the active sheet as UTF-8 CSV using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file (or create a new workbook if it doesn't exist), set the workbook's Title, Author, and Comments properties using reflection for version‑agnostic access, then export the active worksheet to a UTF-8 encoded CSV with a comma separator via TxtSaveOptions. | Write C# code that ensures the output directory exists, configures TxtSaveOptions for CSV output, and saves the workbook after updating its document properties, handling scenarios where WorkbookProperties is unavailable.
// Common Searches: how to set workbook title and author with Aspose.Cells when WorkbookProperties is missing | c# save current sheet as UTF-8 CSV with Aspose.Cells TxtSaveOptions | use reflection to modify Excel document properties in Aspose.Cells .NET | convert Excel workbook to UTF-8 CSV using Aspose.Cells example | create new workbook if input.xlsx not found Aspose.Cells C#
// Tags: Aspose.Cells reflection for workbook properties | TxtSaveOptions CSV export with UTF-8 encoding | fallback when WorkbookProperties not available | auto-create workbook if source file missing | ensure output folder existence in C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The sample loads an existing Excel file or creates a new workbook, uses reflection to set Title, Author, and Comments when WorkbookProperties is present, prepares TxtSaveOptions for UTF-8 CSV output, ensures the target directory exists, and saves the active worksheet as a CSV file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Load workbook if the input file exists; otherwise create a new empty workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Attempt to set document properties using reflection (covers versions where WorkbookProperties may be absent).
            try
            {
                var wpProp = workbook.GetType().GetProperty("WorkbookProperties");
                if (wpProp != null)
                {
                    var wp = wpProp.GetValue(workbook);
                    var titleProp = wp.GetType().GetProperty("Title");
                    var authorProp = wp.GetType().GetProperty("Author");
                    var commentsProp = wp.GetType().GetProperty("Comments");

                    titleProp?.SetValue(wp, "Updated Title");
                    authorProp?.SetValue(wp, "John Doe");
                    commentsProp?.SetValue(wp, "Updated via Aspose.Cells");
                }
                else
                {
                    Console.WriteLine("WorkbookProperties API not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to set workbook properties: {ex.Message}");
            }

            // Configure CSV save options.
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,
                Separator = ','
                // ExportActiveWorksheetOnly is true by default for CSV.
            };

            // Ensure the output directory exists.
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook saved as CSV to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving CSV: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
