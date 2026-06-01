using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsReadNumericAsText
{
    class Program
    {
        static void Main()
        {
            // Sample CSV content with numeric values
            string csvContent = "ID,Price,Quantity\n" +
                                "1,19.99,5\n" +
                                "2,24.50,10\n" +
                                "3,N/A,0";

            // Write the CSV content to a temporary file
            string tempCsvPath = Path.GetTempFileName();
            File.WriteAllText(tempCsvPath, csvContent, Encoding.UTF8);

            // Create TxtLoadOptions and disable numeric conversion
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // When false, numeric strings are kept as text
                ConvertNumericData = false,
                // Keep date conversion disabled for clarity
                ConvertDateTimeData = false,
                // Use comma as separator (default for CSV)
                Separator = ','
            };

            // Load the CSV file with the specified options
            Workbook workbook = new Workbook(tempCsvPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Demonstrate that numeric cells are read as text
            Console.WriteLine("Cell A2 (ID)   Type: " + cells["A2"].Type + "   Value: " + cells["A2"].StringValue);
            Console.WriteLine("Cell B2 (Price) Type: " + cells["B2"].Type + "   Value: " + cells["B2"].StringValue);
            Console.WriteLine("Cell C2 (Qty)   Type: " + cells["C2"].Type + "   Value: " + cells["C2"].StringValue);

            // Save the workbook to an Excel file (optional)
            string outputPath = Path.Combine(Path.GetDirectoryName(tempCsvPath), "Result.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved to: " + outputPath);

            // Clean up temporary CSV file
            File.Delete(tempCsvPath);
        }
    }
}