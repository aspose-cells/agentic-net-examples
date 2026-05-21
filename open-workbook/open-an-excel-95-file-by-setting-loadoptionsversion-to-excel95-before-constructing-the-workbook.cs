using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file (Excel 95 is compatible with Excel 97-2003 format)
        string filePath = "sample.xls";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // LoadOptions can be used to specify the expected format; Excel 95 files are handled as Excel97To2003
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example operation: output the name of the first worksheet
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
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