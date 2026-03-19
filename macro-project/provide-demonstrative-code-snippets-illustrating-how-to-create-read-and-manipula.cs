using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsXpsDemo
{
    public class XpsOperations
    {
        // Creates a new workbook and saves it as XPS (and also as XLSX for later manipulation)
        public static void CreateAndSaveXps()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456.78);

            // Save as XLSX (used for later loading and manipulation)
            workbook.Save("DemoDocument.xlsx");

            // Initialize XpsSaveOptions
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,               // One page per worksheet
                DefaultFont = "Arial",                // Fallback font
                CheckFontCompatibility = true,       // Verify font compatibility
                PageIndex = 0,                        // Start from first page
                PageCount = 1                         // Save only one page (for demo)
            };

            // Save the workbook as XPS
            workbook.Save("DemoDocument.xps", saveOptions);
        }

        // Loads an existing XLSX file into a Workbook object
        public static Workbook LoadWorkbook(string filePath)
        {
            return new Workbook(filePath);
        }

        // Manipulates the loaded workbook and re‑saves it as XPS
        public static void ManipulateAndResaveXps(string sourcePath, string targetPath)
        {
            // Load the existing XLSX document
            Workbook workbook = LoadWorkbook(sourcePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Example manipulation: add a new row with calculated data
            int lastRow = sheet.Cells.MaxDataRow + 1; // zero‑based index of the new row
            sheet.Cells[$"A{lastRow + 1}"].PutValue("Total");
            sheet.Cells[$"B{lastRow + 1}"].Formula = $"SUM(B1:B{lastRow})";

            // Re‑calculate formulas
            workbook.CalculateFormula();

            // Prepare save options
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial",
                CheckFontCompatibility = true,
                PageIndex = 0,
                PageCount = 1
            };

            // Save the modified workbook as XPS
            workbook.Save(targetPath, saveOptions);
        }

        // Entry point for demonstration
        public static void RunDemo()
        {
            // Step 1: Create and save an XPS document (and an XLSX for later use)
            CreateAndSaveXps();

            // Step 2: Load the created XLSX, modify it, and save again as XPS
            ManipulateAndResaveXps("DemoDocument.xlsx", "DemoDocument_Modified.xps");

            Console.WriteLine("XPS creation, loading, and manipulation completed.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            XpsOperations.RunDemo();
        }
    }
}