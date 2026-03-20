using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

class ConsoleExportDemo
{
    static void Main()
    {
        // Create a workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue(999.99);
        worksheet.Cells["A3"].PutValue("Phone");
        worksheet.Cells["B3"].PutValue(599.99);

        // Apply currency format to the price column
        Style currencyStyle = workbook.CreateStyle();
        currencyStyle.Number = 4; // Currency format
        Aspose.Cells.Range priceRange = worksheet.Cells.CreateRange("B2:B3");
        priceRange.SetStyle(currencyStyle);

        // Export the range to a DataTable using DisplayStyle for readable console output
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            FormatStrategy = CellValueFormatStrategy.DisplayStyle,
            ExportColumnName = true
        };
        DataTable table = worksheet.Cells.ExportDataTable(0, 0, 3, 2, exportOptions);
        PrintDataTable(table);

        // Save the worksheet as a tab‑separated TXT file in memory with UTF‑8 encoding
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Encoding = Encoding.UTF8,
            Separator = '\t',
            FormatStrategy = CellValueFormatStrategy.DisplayStyle,
            QuoteType = TxtValueQuoteType.Always
        };
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, txtOptions);
            stream.Position = 0;
            string txtContent = new StreamReader(stream, txtOptions.Encoding).ReadToEnd();

            Console.WriteLine("\n--- TXT Export (UTF‑8, tab‑separated) ---");
            Console.WriteLine(txtContent);
        }
    }

    // Helper method to print a DataTable to the console with tab separation
    static void PrintDataTable(DataTable dt)
    {
        // Print column headers
        foreach (DataColumn col in dt.Columns)
        {
            Console.Write(col.ColumnName + "\t");
        }
        Console.WriteLine();

        // Print each row
        foreach (DataRow row in dt.Rows)
        {
            foreach (var item in row.ItemArray)
            {
                Console.Write(item + "\t");
            }
            Console.WriteLine();
        }
    }
}