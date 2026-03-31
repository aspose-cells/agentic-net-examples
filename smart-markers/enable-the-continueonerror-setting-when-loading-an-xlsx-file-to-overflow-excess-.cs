using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Generate CSV data with more columns than a single worksheet can hold (e.g., 300 columns)
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < 300; i++)
        {
            sb.Append("Col" + i);
            if (i < 299) sb.Append(",");
        }
        // Two rows of data – enough to trigger overflow when loading
        string csvContent = sb.ToString() + "\n" + sb.ToString();

        // Convert the CSV string to a memory stream
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
        {
            // Enable overflow of excess data onto new worksheets
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true
            };

            // Load the CSV data into a workbook using the specified options
            Workbook workbook = new Workbook(csvStream, loadOptions);

            // Save the resulting workbook as XLSX; overflowed data will appear on additional sheets
            workbook.Save("OverflowedData.xlsx");
        }
    }
}