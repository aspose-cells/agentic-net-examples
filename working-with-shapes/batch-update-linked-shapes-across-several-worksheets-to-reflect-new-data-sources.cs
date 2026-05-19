using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace LinkedShapeBatchUpdate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify and load the main workbook
                const string mainPath = "MainWorkbook.xlsx";
                if (!File.Exists(mainPath))
                    throw new FileNotFoundException($"Main workbook not found: {mainPath}");

                Workbook mainWorkbook = new Workbook(mainPath);

                // Prepare external workbooks that exist
                string[] externalPaths = { "DataSource1.xlsx", "DataSource2.xlsx" };
                List<Workbook> externalWorkbooks = new List<Workbook>();

                foreach (string path in externalPaths)
                {
                    if (File.Exists(path))
                    {
                        externalWorkbooks.Add(new Workbook(path));
                    }
                    else
                    {
                        Console.WriteLine($"Warning: External workbook not found and will be skipped: {path}");
                    }
                }

                // Update linked data sources if any external workbooks are available
                if (externalWorkbooks.Count > 0)
                {
                    mainWorkbook.UpdateLinkedDataSource(externalWorkbooks.ToArray());
                }

                // Refresh selected values of all linked shapes in every worksheet
                foreach (Worksheet sheet in mainWorkbook.Worksheets)
                {
                    sheet.Shapes.UpdateSelectedValue();
                }

                // Recalculate formulas to reflect the new data
                mainWorkbook.CalculateFormula();

                // Save the updated workbook
                const string outputPath = "MainWorkbook_Updated.xlsx";
                mainWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}