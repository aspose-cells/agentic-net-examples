// Title: Disable column drag‑and‑drop reordering while allowing column width changes using Aspose.Cells worksheet protection in C#
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet, disables column drag‑and‑drop reordering, insertion, and deletion, but keeps column width formatting enabled. | Show how to set the AllowFormattingColumn property and related protection flags in Aspose.Cells to block column moving while permitting column resizing, then save the workbook.
// Common Searches: Aspose.Cells C# protect worksheet prevent column reordering but allow column resizing | How to disable column drag and drop in an Excel sheet using Aspose.Cells | C# Aspose.Cells worksheet protection settings for column operations | Allow column width changes while locking column order with Aspose.Cells | Set worksheet protection without password in Aspose.Cells C#
// Tags: Aspose.Cells worksheet protection column reordering | C# allow column width formatting Aspose.Cells | disable column insertion deletion Aspose.Cells | protect Excel sheet without password Aspose.Cells | AllowFormattingColumn property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, accesses the first worksheet, enables column width formatting while disabling column insertion, deletion, row formatting, and sorting, applies full protection without a password, and saves the result as ProtectedSheet.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set protection options (use singular property names)
                worksheet.Protection.AllowFormattingColumn = true;
                worksheet.Protection.AllowInsertingColumn = false;
                worksheet.Protection.AllowDeletingColumn = false;
                worksheet.Protection.AllowFormattingRow = false;
                worksheet.Protection.AllowSorting = false;
                // The AllowAutoFilter property is not available in this version of Aspose.Cells; omitted.

                // Apply protection (no password)
                worksheet.Protect(ProtectionType.All);

                // Save the protected workbook
                string outputPath = "ProtectedSheet.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }
}
