using System;
using System.IO;
using Aspose.Cells;

class VlookupExternalDemo
{
    static void Main()
    {
        try
        {
            // ---------- Create and save the external workbook ----------
            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "LookupData";

            // Header
            extSheet.Cells["A1"].PutValue("Key");
            extSheet.Cells["B1"].PutValue("Value");

            // Sample lookup table
            extSheet.Cells["A2"].PutValue("Apple");
            extSheet.Cells["B2"].PutValue(10);
            extSheet.Cells["A3"].PutValue("Banana");
            extSheet.Cells["B3"].PutValue(20);

            // Ensure folder exists and save the external workbook
            string folder = Path.GetFullPath("ExternalFiles");
            Directory.CreateDirectory(folder);
            string externalPath = Path.Combine(folder, "LookupSource.xlsx");
            externalWb.Save(externalPath);

            // Verify that the external file was created
            if (!File.Exists(externalPath))
                throw new FileNotFoundException("External workbook was not saved.", externalPath);

            // ---------- Create the main workbook ----------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "Main";

            // Value to look up
            mainSheet.Cells["A2"].PutValue("Banana");

            // VLOOKUP formula referencing the external workbook.
            // Note: Use single quotes around the full path, then sheet name, then range.
            string formula = $"=VLOOKUP(A2,'{externalPath}'!LookupData!$A$2:$B$3,2,FALSE)";

            // Set the formula (use the Formula property)
            mainSheet.Cells["B2"].Formula = formula;

            // ---------- Calculate the formula using the external workbook as a linked data source ----------
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Provide the external workbook as a linked data source
                LinkedDataSources = new Workbook[] { externalWb }
            };
            mainWb.CalculateFormula(calcOptions);

            // Display the result of the VLOOKUP
            Console.WriteLine("Lookup result: " + mainSheet.Cells["B2"].Value);

            // ---------- Save the main workbook ----------
            string mainPath = Path.Combine(folder, "MainWorkbook.xlsx");
            mainWb.Save(mainPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}