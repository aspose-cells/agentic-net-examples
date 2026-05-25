using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create an external workbook ----------
                Workbook externalWb = new Workbook();
                Worksheet extSheet = externalWb.Worksheets[0];
                extSheet.Name = "Sheet1";
                // Put a value that will be referenced from the main workbook
                extSheet.Cells["A1"].PutValue(12345);
                // Save the external workbook to a file
                string externalFile = "ExternalData.xlsx";
                externalWb.Save(externalFile, SaveFormat.Xlsx);

                // Ensure the external file exists before creating a link
                if (!File.Exists(externalFile))
                    throw new FileNotFoundException($"External workbook not found: {externalFile}");

                // ---------- Create the main workbook ----------
                Workbook mainWb = new Workbook();
                Worksheet mainSheet = mainWb.Worksheets[0];
                mainSheet.Name = "MainSheet";

                // Add an external link to the external workbook
                int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFile, new[] { "Sheet1" });
                ExternalLink extLink = mainWb.Worksheets.ExternalLinks[linkIndex];

                // ---------- Set a formula that references the external workbook ----------
                // The second argument indicates that this is not an array formula
                Cell targetCell = mainSheet.Cells["B2"];
                targetCell.SetFormula($"=[{externalFile}]Sheet1!A1", false);

                // ---------- Update the external data source and calculate ----------
                mainWb.UpdateLinkedDataSource(new[] { externalWb });
                mainWb.CalculateFormula();

                // ---------- Output the result ----------
                Console.WriteLine("Value in B2 (referencing external workbook): " + targetCell.Value);

                // ---------- Save the main workbook ----------
                mainWb.Save("MainWorkbookWithExternalFormula.xlsx", SaveFormat.Xlsx);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine("File error: " + fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}