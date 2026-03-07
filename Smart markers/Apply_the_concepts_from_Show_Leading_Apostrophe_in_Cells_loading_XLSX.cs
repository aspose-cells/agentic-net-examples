using System;
using Aspose.Cells;

namespace ShowLeadingApostropheDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook with a leading apostrophe ----------
            Workbook createWb = new Workbook();                     // create workbook
            Worksheet createWs = createWb.Worksheets[0];            // get first worksheet

            // Put a value that starts with a single quote
            Cell createCell = createWs.Cells["B10"];
            createCell.PutValue("'12345"); // the leading apostrophe is a formatting marker

            // Enable the QuotePrefix style so the apostrophe is treated as a prefix
            Style style = createCell.GetStyle();
            style.QuotePrefix = true;
            createCell.SetStyle(style);

            // Save the workbook to a file
            string filePath = "LeadingApostrophe.xlsx";
            createWb.Save(filePath);                               // save workbook

            // ---------- Load the workbook and verify the QuotePrefix ----------
            Workbook loadWb = new Workbook(filePath);               // load workbook
            Worksheet loadWs = loadWb.Worksheets[0];
            Cell loadedCell = loadWs.Cells["B10"];

            // Retrieve the QuotePrefix property
            bool isQuotePrefixSet = loadedCell.GetStyle().QuotePrefix;

            // Output the result
            Console.WriteLine($"Cell B10 QuotePrefix is set: {isQuotePrefixSet}");
            Console.WriteLine($"Cell B10 displayed value: {loadedCell.StringValue}");
        }
    }
}