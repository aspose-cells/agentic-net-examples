using System;
using Aspose.Cells;

class OdsPasswordProtection
{
    static void Main()
    {
        // Input XLSX file path
        string inputFile = "input.xlsx";

        // Output ODS file path
        string outputFile = "protected_output.ods";

        // Password to protect the workbook
        string password = "mySecret";

        // Load the source XLSX workbook
        Workbook workbook = new Workbook(inputFile);

        // Apply password protection (encryption) to the workbook
        workbook.Settings.Password = password;

        // Save the workbook as ODS with the applied password
        workbook.Save(outputFile, SaveFormat.Ods);

        // Verify that the ODS file is password protected by loading it with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook protectedWorkbook = new Workbook(outputFile, loadOptions);

        // Output a cell value to confirm successful load
        Console.WriteLine("Cell A1 value after loading protected ODS: " + protectedWorkbook.Worksheets[0].Cells["A1"].Value);
    }
}