using Aspose.Cells;
using System;
using System.IO;

class UnprotectWorksheetDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "ProtectedWorkbook.xlsx";
            const string outputPath = "UnprotectedWorkbook.xlsx";
            const string workbookPassword = "password123"; // password for the workbook file

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with the required password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet (same password as used for the workbook)
            worksheet.Unprotect(workbookPassword);

            // Confirm that the worksheet is no longer protected
            Console.WriteLine("Worksheet is protected: " + worksheet.IsProtected);

            // Save the unprotected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}