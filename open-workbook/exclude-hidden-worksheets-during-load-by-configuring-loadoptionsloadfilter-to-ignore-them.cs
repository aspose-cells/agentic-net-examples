using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenSheetFilterDemo
{
    // Custom filter that loads only visible worksheets
    public class VisibleSheetLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load the sheet only if it is visible
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path for the temporary workbook
            string filePath = "HiddenSheetFilterDemo.xlsx";

            try
            {
                // ---------- Create a workbook with hidden and visible sheets ----------
                using (Workbook wb = new Workbook())
                {
                    // Visible sheet (add first to avoid having zero visible sheets)
                    Worksheet visibleSheet = wb.Worksheets[0];
                    visibleSheet.Name = "VisibleSheet";
                    visibleSheet.Cells["A1"].PutValue("This will be loaded");

                    // Add a hidden sheet
                    Worksheet hiddenSheet = wb.Worksheets.Add("HiddenSheet");
                    hiddenSheet.IsVisible = false;
                    hiddenSheet.Cells["A1"].PutValue("This should not be loaded");

                    // Save the workbook
                    wb.Save(filePath);
                }

                // ---------- Load the workbook using a custom LoadFilter ----------
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new VisibleSheetLoadFilter()
                };

                using (Workbook loadedWb = new Workbook(filePath, loadOptions))
                {
                    // Verify that only the visible sheet was loaded
                    Console.WriteLine("Loaded worksheets:");
                    foreach (Worksheet ws in loadedWb.Worksheets)
                    {
                        Console.WriteLine($"- {ws.Name} (Visible={ws.IsVisible})");
                        Console.WriteLine($"  A1 Value: {ws.Cells["A1"].StringValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}