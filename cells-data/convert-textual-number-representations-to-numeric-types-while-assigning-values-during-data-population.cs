// Title: Convert string numbers and dates to native types in Aspose.Cells for .NET (C#)
// Description: This example shows how to populate cells with numeric, date and non‑numeric strings, invoke Cells.ConvertStringToNumericValue to transform convertible strings into their CLR types, and verify the results. It also demonstrates loading CSV data from a memory stream with TxtLoadOptions (ConvertNumericData and ConvertDateTimeData enabled) so that numeric and date strings are automatically converted, then saves both workbooks.
// Keywords: Aspose.Cells ConvertStringToNumericValue | C# string to numeric conversion | Aspose.Cells CSV load options | ConvertDateTimeData Aspose.Cells | numeric string to double Aspose.Cells | date string to DateTime Aspose.Cells
// Common Searches: Aspose.Cells convert string to numeric value | How to auto‑convert dates when loading CSV in Aspose.Cells | C# Cells.ConvertStringToNumericValue example | TxtLoadOptions ConvertNumericData true | ConvertDateTimeData option Aspose.Cells
// Developer Intent: Transform textual representations of numbers and dates into native numeric and DateTime objects while populating a workbook or importing CSV data with Aspose.Cells for .NET.
// Use Cases: Insert numeric and date strings into cells, call Cells.ConvertStringToNumericValue, and read the CLR type of each cell. | Import a CSV stream using TxtLoadOptions with ConvertNumericData and ConvertDateTimeData set to true, so values are stored as double and DateTime automatically. | Validate conversion by checking cell.Value.GetType(), then save the workbook for further processing.
// AI Prompts: Show C# code that uses Cells.ConvertStringToNumericValue to change string numbers and dates into native types in an Aspose.Cells workbook. | Demonstrate loading a CSV file with Aspose.Cells TxtLoadOptions so numeric and date strings are automatically converted. | Write a script that logs the CLR type of each cell after conversion and saves the workbook as an Excel file.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// This example shows how to populate cells with numeric, date and non‑numeric strings, invoke Cells.ConvertStringToNumericValue to transform convertible strings into their CLR types, and verify the results. It also demonstrates loading CSV data from a memory stream with TxtLoadOptions (ConvertNumericData and ConvertDateTimeData enabled) so that numeric and date strings are automatically converted, then saves both workbooks.
class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook wb = new Workbook();                     // create rule
        Cells cells = wb.Worksheets[0].Cells;

        // ---------- Populate cells with textual representations ----------
        cells["A1"].PutValue("123.45");                  // numeric string
        cells["B1"].PutValue("2021-06-20");              // date string
        cells["C1"].PutValue("NotANumber");              // non‑numeric string

        // ---------- Convert convertible strings to native types ----------
        cells.ConvertStringToNumericValue();              // rule: Cells.ConvertStringToNumericValue

        // ---------- Display the converted values and their CLR types ----------
        Console.WriteLine($"A1: {cells["A1"].Value} (type {cells["A1"].Value.GetType()})");
        Console.WriteLine($"B1: {cells["B1"].Value} (type {cells["B1"].Value.GetType()})");
        Console.WriteLine($"C1: {cells["C1"].Value} (type {cells["C1"].Value.GetType()})");

        // ---------- Load CSV data with automatic conversion ----------
        string csvData = "ID,Amount,Date\n1,99.99,2023-01-15\n2,abc,2023-02-20";
        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            TxtLoadOptions txtOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                ConvertNumericData = true,               // convert numeric strings
                ConvertDateTimeData = true               // convert date strings
            };

            Workbook csvWb = new Workbook(ms, txtOptions); // load rule
            Worksheet ws = csvWb.Worksheets[0];

            // Show conversion results for CSV cells
            Console.WriteLine($"CSV B2 (Amount) type: {ws.Cells["B2"].Value.GetType()}, value: {ws.Cells["B2"].Value}");
            Console.WriteLine($"CSV C2 (Date) type: {ws.Cells["C2"].Value.GetType()}, value: {ws.Cells["C2"].Value}");

            // Save the CSV‑derived workbook
            csvWb.Save("CsvConverted.xlsx");               // save rule
        }

        // ---------- Save the original workbook ----------
        wb.Save("StringConverted.xlsx");                  // save rule
    }
}
