using System;
using Aspose.Cells;

class MergeWorkbooksDemo
{
    static void Main()
    {
        // Paths of the workbooks to be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
        // Path for the merged workbook
        string destFile = "MergedWorkbook.xlsx";

        // Create an empty destination workbook
        Workbook destWorkbook = new Workbook();

        // Enable compatibility check for older Excel versions
        destWorkbook.Settings.CheckCompatibility = true;

        // Load each source workbook, validate data, and combine into destination
        foreach (string srcPath in sourceFiles)
        {
            // Load options with data validation checking enabled
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CheckDataValid = true;

            // Load the source workbook using the load options
            Workbook srcWorkbook = new Workbook(srcPath, loadOptions);

            // Ensure the source workbook also respects compatibility settings
            srcWorkbook.Settings.CheckCompatibility = true;

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(srcWorkbook);
        }

        // Save options to validate merged cells and merge conditional formatting/validation areas
        XlsSaveOptions saveOptions = new XlsSaveOptions
        {
            ValidateMergedAreas = true,
            MergeAreas = true
        };

        // Save the merged workbook with the specified save options
        destWorkbook.Save(destFile, saveOptions);
    }
}