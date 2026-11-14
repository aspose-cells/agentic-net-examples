// Title: Remove every slicer from an Aspose.Cells worksheet and save the workbook as XLSX using C#
// Common Searches: C# code to delete all slicers from an Aspose.Cells worksheet | Aspose.Cells remove slicer collection before saving workbook | How to clear slicers in a .NET workbook and export to XLSX | Worksheet.Slicers.Clear vs RemoveAt loop Aspose.Cells example | Programmatically erase slicers in Excel file with Aspose.Cells for .NET
// Tags: remove all slicers Aspose.Cells | worksheet slicer collection clear C# | export workbook to XLSX after slicer removal | Aspose.Cells slicer management API | clear slicer collection before saving Excel file

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates or loads a workbook, accesses the first worksheet, checks for existing slicers, clears the slicer collection using Clear() or a backward RemoveAt loop, and then saves the workbook as "ClearedSlicers.xlsx" in XLSX format.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Remove all slicers from the worksheet safely
                if (worksheet.Slicers != null && worksheet.Slicers.Count > 0)
                {
                    // Attempt to clear the collection; if not available, remove items individually
                    try
                    {
                        worksheet.Slicers.Clear();
                    }
                    catch (MissingMethodException)
                    {
                        for (int i = worksheet.Slicers.Count - 1; i >= 0; i--)
                        {
                            worksheet.Slicers.RemoveAt(i);
                        }
                    }
                }

                // Save the workbook
                string outputPath = "ClearedSlicers.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
