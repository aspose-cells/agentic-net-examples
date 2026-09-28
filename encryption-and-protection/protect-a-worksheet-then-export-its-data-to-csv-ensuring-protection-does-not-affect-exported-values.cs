// Title: How to password‑protect an Excel worksheet with Aspose.Cells for .NET and export it to CSV without losing data
// AI Prompts: Write C# code that uses Aspose.Cells to apply a password to a worksheet and then saves the workbook as a CSV file. | Show how to enable full worksheet protection in Aspose.Cells and export the protected sheet to CSV while preserving all cell values. | Provide a C# example that creates a workbook, protects the first sheet with ProtectionType.All, creates the output folder if needed, and writes the sheet to a CSV file.
// Common Searches: aspnet protect excel worksheet password aspocells export csv | does worksheet protection affect csv export in Aspose.Cells | c# Aspose.Cells save protected sheet as csv | how to use TxtSaveOptions to export protected worksheet to csv | protect excel sheet with Aspose.Cells then convert to csv
// Tags: Aspose.Cells worksheet password protection C# | CSV export of protected worksheet Aspose.Cells | TxtSaveOptions for CSV Aspose.Cells | using ProtectionType.All in Aspose.Cells | pre‑create output folder Aspose.Cells CSV

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data, applies full password protection to the first worksheet using ProtectionType.All, ensures the target directory exists, and then saves the workbook as a CSV file via TxtSaveOptions, demonstrating that protection does not alter the exported values.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Protect the worksheet (all protection types) with a password.
            // The third parameter is the old password; an empty string is used when there is no previous password.
            sheet.Protect(ProtectionType.All, "myPassword", string.Empty);

            // Export the worksheet to CSV.
            // Worksheet protection does not affect the data exported to CSV.
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            string outputPath = "ExportedData.csv";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
