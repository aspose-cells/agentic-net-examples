using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCloneNamedRanges
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and add named ranges ----------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                sourceSheet.Name = "Data";

                // Populate some data
                sourceSheet.Cells["A1"].PutValue("Item");
                sourceSheet.Cells["B1"].PutValue("Quantity");
                sourceSheet.Cells["A2"].PutValue("Apple");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("Banana");
                sourceSheet.Cells["B3"].PutValue(20);

                // Create first named range
                int idx1 = sourceWorkbook.Worksheets.Names.Add("ItemRange");
                sourceWorkbook.Worksheets.Names[idx1].RefersTo = "=Data!$A$2:$A$3";

                // Create second named range
                int idx2 = sourceWorkbook.Worksheets.Names.Add("QtyRange");
                sourceWorkbook.Worksheets.Names[idx2].RefersTo = "=Data!$B$2:$B$3";

                // ---------- Clone the workbook preserving named ranges ----------
                Workbook clonedWorkbook = new Workbook(); // empty destination workbook
                CopyOptions options = new CopyOptions
                {
                    CopyNames = true // ensure named ranges are copied
                };
                clonedWorkbook.Copy(sourceWorkbook, options);

                // ---------- Verify that all named ranges are present in the clone ----------
                Console.WriteLine("Verifying named ranges in the cloned workbook:");
                foreach (Name name in clonedWorkbook.Worksheets.Names)
                {
                    // Retrieve the range the name refers to
                    AsposeRange range = name.GetRange();
                    Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}, Address: {range?.Address}");
                }

                // Optional: Save both workbooks for visual inspection
                string sourcePath = "SourceWorkbook.xlsx";
                string clonedPath = "ClonedWorkbook.xlsx";

                // Ensure we can write to the target locations
                try
                {
                    sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);
                    clonedWorkbook.Save(clonedPath, SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbooks: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}