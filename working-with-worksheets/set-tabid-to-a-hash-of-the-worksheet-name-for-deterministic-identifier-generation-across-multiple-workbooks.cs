// Title: Assign a deterministic Worksheet.TabId from the sheet name using SHA‑256 in Aspose.Cells for .NET
// Common Searches: how can I generate a consistent TabId for Excel worksheets with Aspose.Cells C# | set worksheet TabId based on sheet name hash in Aspose.Cells .NET | deterministic TabId assignment for multiple workbooks using Aspose.Cells | compute non‑negative TabId from worksheet name in C# Aspose.Cells example | persist custom TabId values when saving and reloading a workbook with Aspose.Cells
// Tags: deterministic worksheet tabid generation c# | sha256 hash based tabid aspose.cells | assign tabid from sheet name aspose.cells | nonnegative tabid calculation .net | persist custom tabid across workbooks aspose

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The sample defines a ComputeTabId method that creates a SHA‑256 hash of a worksheet's name, extracts the first four bytes as a little‑endian integer, masks the sign bit to ensure a non‑negative value, and assigns this integer to each sheet's TabId. It loads or creates a workbook, applies the deterministic TabId to all worksheets, saves the file, and reloads it to verify that the custom TabId values are persisted.
class Program
{
    /// <param name="sheetName">The worksheet name.</param>
    /// <returns>An int suitable for Worksheet.TabId.</returns>
    private static int ComputeTabId(string sheetName)
    {
        if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));

        using (SHA256 sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(sheetName));
            // Take first 4 bytes (little‑endian) and convert to int
            int id = BitConverter.ToInt32(hash, 0);
            // Ensure non‑negative (TabId expects a non‑negative value)
            return id & 0x7FFFFFFF;
        }
    }

    static void Main(string[] args)
    {
        try
        {
            // Determine input and output file paths
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            Workbook workbook;

            if (File.Exists(inputPath))
            {
                Console.WriteLine($"Loading workbook from '{inputPath}'...");
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file not found. Creating a new workbook with sample sheets.");
                workbook = new Workbook();

                // Add a few sample worksheets
                workbook.Worksheets[0].Name = "Summary";
                workbook.Worksheets.Add("Data");
                workbook.Worksheets.Add("Report");
            }

            // Iterate through all worksheets and set deterministic TabId
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int deterministicId = ComputeTabId(sheet.Name);
                sheet.TabId = deterministicId;
                Console.WriteLine($"Worksheet '{sheet.Name}' assigned TabId = {deterministicId}");
            }

            // Save the workbook
            Console.WriteLine($"Saving workbook to '{outputPath}'...");
            workbook.Save(outputPath);
            Console.WriteLine("Save completed.");

            // Reload to verify TabId values persisted
            Console.WriteLine("Reloading saved workbook for verification...");
            Workbook verifyWorkbook = new Workbook(outputPath);
            Console.WriteLine("Verification results:");
            foreach (Worksheet sheet in verifyWorkbook.Worksheets)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' has TabId = {sheet.TabId}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred:");
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
