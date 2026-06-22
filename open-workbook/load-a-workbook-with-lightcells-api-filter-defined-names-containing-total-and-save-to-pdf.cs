using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsLightCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Create LoadOptions and set a LoadFilter that loads only defined names
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

            // Load the workbook using the LightCells API (LoadOptions)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Collect names that do NOT contain "Total" so they can be removed
            List<string> namesToRemove = new List<string>();
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Name.Text holds the name string (e.g., "TotalSales")
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(definedName.Text);
                }
            }

            // Remove the unwanted defined names from the collection
            foreach (string name in namesToRemove)
            {
                // Remove by name; this method exists in NameCollection
                workbook.Worksheets.Names.Remove(name);
            }

            // Save the filtered workbook as PDF using the provided Save method
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook loaded, filtered, and saved to PDF at: {pdfPath}");
        }
    }
}