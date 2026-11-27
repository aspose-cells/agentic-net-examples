// Title: Load an Excel workbook from a MemoryStream, modify slicer selections, refresh pivot tables, and save as PDF with Aspose.Cells for .NET
// Common Searches: c# aspnet load xlsx from memory stream modify slicer cache items | asp.net aspose.cells update slicer selection programmatically | refresh pivot tables after slicer change using Aspose.Cells C# | export workbook to PDF directly from MemoryStream with Aspose.Cells | how to select a single slicer item in a loaded workbook using Aspose.Cells
// Tags: slicer cache item selection Aspose.Cells | refresh pivot tables Aspose.Cells C# | save workbook as PDF from MemoryStream Aspose.Cells | load workbook from MemoryStream Aspose.Cells | export Excel to PDF with slicer state Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfDemo
{
    // The example creates a workbook with data, a pivot table and a slicer, saves it to a MemoryStream, reloads it, programmatically selects only the first slicer item, refreshes all pivot tables to apply the slicer change, and finally writes the workbook as a PDF into another MemoryStream using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook with data, a pivot table and a slicer
                // ------------------------------------------------------------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sheet = sourceWorkbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                // Header
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                // Data rows
                string[,] data = new string[,]
                {
                    { "Apple",  "2020", "50" },
                    { "Banana", "2020", "60" },
                    { "Apple",  "2021", "70" },
                    { "Banana", "2021", "80" },
                    { "Cherry", "2020", "90" },
                    { "Cherry", "2021", "100" }
                };

                for (int r = 0; r < data.GetLength(0); r++)
                {
                    cells[r + 1, 0].PutValue(data[r, 0]);
                    cells[r + 1, 1].PutValue(int.Parse(data[r, 1]));
                    cells[r + 1, 2].PutValue(double.Parse(data[r, 2]));
                }

                // Create a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:C7", "E3", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the "Fruit" field of the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "Fruit", "F2");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // ------------------------------------------------------------
                // 2. Save the workbook to a memory stream (XLSX format)
                // ------------------------------------------------------------
                using (MemoryStream workbookStream = new MemoryStream())
                {
                    sourceWorkbook.Save(workbookStream, SaveFormat.Xlsx);
                    workbookStream.Position = 0; // Reset for reading

                    // ------------------------------------------------------------
                    // 3. Load the workbook from the memory stream
                    // ------------------------------------------------------------
                    Workbook loadedWorkbook = new Workbook(workbookStream);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                    // ------------------------------------------------------------
                    // 4. Update slicer selections
                    // ------------------------------------------------------------
                    // For demonstration, select the first item and deselect others
                    foreach (Slicer loadedSlicer in loadedSheet.Slicers)
                    {
                        SlicerCache cache = loadedSlicer.SlicerCache;
                        SlicerCacheItemCollection items = cache.SlicerCacheItems;

                        Console.WriteLine($"Slicer '{loadedSlicer.Name}' has {items.Count} items.");

                        for (int i = 0; i < items.Count; i++)
                        {
                            // Select only the first item (index 0)
                            items[i].Selected = (i == 0);
                            Console.WriteLine($"  Item {i}: Value='{items[i].Value}', Selected={items[i].Selected}");
                        }
                    }

                    // ------------------------------------------------------------
                    // 5. Refresh all pivot tables (ensures they reflect slicer changes)
                    // ------------------------------------------------------------
                    loadedWorkbook.Worksheets.RefreshAll();

                    // ------------------------------------------------------------
                    // 6. Save the workbook as PDF into another memory stream
                    // ------------------------------------------------------------
                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        // Optional: customize PDF save options
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            // Example options – adjust as needed
                            OnePagePerSheet = true,
                            Compliance = PdfCompliance.PdfA1b
                        };

                        loadedWorkbook.Save(pdfStream, pdfOptions);
                        Console.WriteLine($"PDF generated. Stream length = {pdfStream.Length} bytes.");

                        // For verification you may write the PDF to a file (optional)
                        // File.WriteAllBytes("Result.pdf", pdfStream.ToArray());
                    }
                }

                Console.WriteLine("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
