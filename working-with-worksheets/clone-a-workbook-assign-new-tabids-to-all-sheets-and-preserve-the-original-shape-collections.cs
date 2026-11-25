// Title: Clone an Excel workbook, assign new TabIds to each worksheet, and keep chart shapes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy workbook and change worksheet TabId in C# | preserve chart shapes when cloning an Excel file with Aspose.Cells | verify shape collection count after cloning a workbook using Aspose.Cells | assign sequential TabId values to cloned worksheets in .NET | deep copy Excel workbook while retaining drawings with Aspose.Cells
// Tags: Workbook.Copy deep copy Aspose.Cells | set Worksheet.TabId programmatically | retain chart objects during workbook clone | compare shape counts between source and clone | clone Excel workbook preserving drawings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WorkbookCloneDemo
{
    // The example loads or creates a sample Excel file with two worksheets that contain chart shapes, clones the workbook using the Workbook.Copy method, assigns new TabId values to each cloned worksheet starting from 1000, validates that the shape collections (charts) are preserved by comparing counts with the source sheets, saves the cloned workbook, and prints detailed worksheet information for both source and cloned workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the source and cloned workbooks
            string sourcePath = "source.xlsx";
            string clonedPath = "cloned.xlsx";

            try
            {
                // Ensure the source workbook exists; if not, create a sample workbook with shapes
                if (!File.Exists(sourcePath))
                {
                    try
                    {
                        CreateSampleWorkbook(sourcePath);
                        Console.WriteLine($"Sample source workbook created at '{sourcePath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                        return;
                    }
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Console.WriteLine("Source workbook loaded.");
                PrintWorkbookInfo(sourceWorkbook, "Source");

                // Clone the workbook (deep copy) using Copy method
                Workbook clonedWorkbook = new Workbook();
                clonedWorkbook.Copy(sourceWorkbook);
                Console.WriteLine("Workbook cloned.");

                // Assign new TabIds and validate shape collections
                AssignNewTabIdsAndValidateShapes(sourceWorkbook, clonedWorkbook);

                // Save the cloned workbook
                clonedWorkbook.Save(clonedPath);
                Console.WriteLine($"Cloned workbook saved to '{clonedPath}'.");

                // Print information about the cloned workbook
                PrintWorkbookInfo(clonedWorkbook, "Cloned");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();

            // First sheet data
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "DataSheet1";
            ws1.Cells["A1"].PutValue("Category");
            ws1.Cells["B1"].PutValue("Value");
            ws1.Cells["A2"].PutValue("A");
            ws1.Cells["B2"].PutValue(10);
            ws1.Cells["A3"].PutValue("B");
            ws1.Cells["B3"].PutValue(20);
            ws1.Cells["A4"].PutValue("C");
            ws1.Cells["B4"].PutValue(30);

            // Add a chart shape to first sheet
            int chartIndex1 = ws1.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart1 = ws1.Charts[chartIndex1];
            chart1.NSeries.Add("A1:B4", true);
            chart1.Title.Text = "Sample Chart 1";

            // Second sheet (Add returns Worksheet)
            Worksheet ws2 = wb.Worksheets.Add("DataSheet2");
            ws2.Cells["A1"].PutValue("X");
            ws2.Cells["B1"].PutValue("Y");
            ws2.Cells["A2"].PutValue(1);
            ws2.Cells["B2"].PutValue(5);
            ws2.Cells["A3"].PutValue(2);
            ws2.Cells["B3"].PutValue(15);
            ws2.Cells["A4"].PutValue(3);
            ws2.Cells["B4"].PutValue(25);

            // Add a chart shape to second sheet
            int chartIndex2 = ws2.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart2 = ws2.Charts[chartIndex2];
            chart2.NSeries.Add("A1:B4", true);
            chart2.Title.Text = "Sample Chart 2";

            wb.Save(path);
        }

        private static void AssignNewTabIdsAndValidateShapes(Workbook source, Workbook clone)
        {
            int nextTabId = 1000;

            foreach (Worksheet srcSheet in source.Worksheets)
            {
                Worksheet clonedSheet = clone.Worksheets[srcSheet.Index];

                // Assign a new TabId
                clonedSheet.TabId = nextTabId++;
                Console.WriteLine($"Worksheet '{clonedSheet.Name}' assigned new TabId: {clonedSheet.TabId}");

                // Validate shape collection count
                int srcShapeCount = srcSheet.Shapes.Count;
                int clonedShapeCount = clonedSheet.Shapes.Count;

                if (srcShapeCount != clonedShapeCount)
                {
                    Console.WriteLine($"Warning: Shape count mismatch on sheet '{srcSheet.Name}'. Source={srcShapeCount}, Cloned={clonedShapeCount}");
                }
                else
                {
                    Console.WriteLine($"Shape collection preserved for sheet '{srcSheet.Name}' (Count={srcShapeCount}).");
                }
            }
        }

        private static void PrintWorkbookInfo(Workbook wb, string label)
        {
            Console.WriteLine($"--- {label} Workbook Info ---");
            Console.WriteLine($"Total Worksheets: {wb.Worksheets.Count}");
            foreach (Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine($"Sheet Index: {ws.Index}, Name: '{ws.Name}', TabId: {ws.TabId}, Shapes: {ws.Shapes.Count}");
            }
            Console.WriteLine($"--- End of {label} Info ---");
        }
    }
}
