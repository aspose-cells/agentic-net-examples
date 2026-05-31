using System;
using Aspose.Cells;

namespace AsposeCellsCrossFileNamedRange
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create an external workbook that will hold the source data
            // -----------------------------------------------------------------
            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Sheet1";

            // Populate some data in the external workbook (A1:A3)
            extSheet.Cells["A1"].PutValue(10);
            extSheet.Cells["A2"].PutValue(20);
            extSheet.Cells["A3"].PutValue(30);

            // Save the external workbook to disk
            string externalFileName = "ExternalData.xlsx";
            externalWb.Save(externalFileName);

            // ---------------------------------------------------------------
            // Step 2: Create the main workbook where the named range will live
            // ---------------------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "MainSheet";

            // ---------------------------------------------------------------
            // Step 3: Add an external link to the main workbook pointing to the file created above
            // ---------------------------------------------------------------
            // The external link collection requires the file name and the sheet names it references
            string[] externalSheetNames = new string[] { "Sheet1" };
            int externalLinkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFileName, externalSheetNames);
            ExternalLink externalLink = mainWb.Worksheets.ExternalLinks[externalLinkIndex];

            // (Optional) Add an external name that points to a range in the external workbook
            // This allows using a simpler reference like =[ExternalData.xlsx]!ExtRange
            externalLink.AddExternalName("ExtRange", "=Sheet1!$A$1:$A$3");

            // ---------------------------------------------------------------
            // Step 4: Create a named range in the main workbook that references the external range
            // ---------------------------------------------------------------
            // Add a new name to the workbook's Names collection
            int nameIndex = mainWb.Worksheets.Names.Add("CrossFileRange");
            Name crossFileName = mainWb.Worksheets.Names[nameIndex];

            // Set the RefersTo property to point to the external range.
            // You can reference directly using the external file name and sheet,
            // or use the external name defined above.
            // Example using direct reference:
            // crossFileName.RefersTo = "=[ExternalData.xlsx]Sheet1!$A$1:$A$3";

            // Example using the external name (uncomment the line below if you prefer this style):
            // crossFileName.RefersTo = "=[ExternalData.xlsx]!ExtRange";

            // Here we use the direct reference:
            crossFileName.RefersTo = "=[ExternalData.xlsx]Sheet1!$A$1:$A$3";

            // ---------------------------------------------------------------
            // Step 5: Use the named range in a formula inside the main workbook
            // ---------------------------------------------------------------
            // Place a formula that sums the external range
            mainSheet.Cells["B1"].Formula = "=SUM(CrossFileRange)";

            // Calculate all formulas in the main workbook
            mainWb.CalculateFormula();

            // Output the result of the formula to the console
            Console.WriteLine("Result of SUM(CrossFileRange): " + mainSheet.Cells["B1"].Value);

            // ---------------------------------------------------------------
            // Step 6: Save the main workbook
            // ---------------------------------------------------------------
            string mainFileName = "MainWorkbook.xlsx";
            mainWb.Save(mainFileName);

            Console.WriteLine($"Main workbook saved as '{mainFileName}'.");
        }
    }
}