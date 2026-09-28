// Title: Generate an Excel column chart from XML‑mapped categories and amounts and embed it in a worksheet using Aspose.Cells for C#
// AI Prompts: Write C# code that parses an XML string containing Category and Amount elements, writes the data to an Aspose.Cells worksheet, adds a SUM total row, and inserts a column chart below the data. | Show how to use MaxDataRow to calculate a dynamic range for the chart series and set the chart title in an Aspose.Cells workbook. | Demonstrate saving the workbook as an .xlsx file after populating the sheet and embedding the chart.
// Common Searches: how to create a column chart from XML data using Aspose.Cells in C# | c# Aspose.Cells add total row with SUM formula based on XML‑mapped cells | dynamic chart range using MaxDataRow in Aspose.Cells workbook | embed chart below data rows in Excel file generated with Aspose.Cells C#
// Tags: Aspose.Cells create column chart from XML data | C# parse XML to populate Excel worksheet | Aspose.Cells add SUM total row | dynamic chart series range MaxDataRow Aspose.Cells | save workbook as .xlsx Aspose.Cells C#

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, parses a hard‑coded XML string with Category and Amount elements, writes the values to columns A and B, adds header cells, inserts a total label with a SUM formula for the Amount column, creates a column chart positioned below the total row that references the populated data range, and saves the file as ChartFromXml.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Sample XML data
            string xmlData = @"<Root>
    <Item>
        <Category>Food</Category>
        <Amount>120</Amount>
    </Item>
    <Item>
        <Category>Transport</Category>
        <Amount>80</Amount>
    </Item>
    <Item>
        <Category>Utilities</Category>
        <Amount>150</Amount>
    </Item>
</Root>";

            // Parse XML and write data starting at row 2 (A2, B2, ...)
            XDocument doc = XDocument.Parse(xmlData);
            int currentRow = 1; // zero‑based index; row 2 in Excel
            foreach (var item in doc.Root.Elements("Item"))
            {
                string category = item.Element("Category")?.Value ?? string.Empty;
                string amountStr = item.Element("Amount")?.Value ?? "0";

                sheet.Cells[currentRow, 0].PutValue(category);
                sheet.Cells[currentRow, 1].PutValue(double.Parse(amountStr));
                currentRow++;
            }

            // Add column headers
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");

            // Determine the last row that contains data (zero‑based)
            int lastDataRow = sheet.Cells.MaxDataRow; // e.g., 3 for three items (rows 2‑4)

            // Insert a total label and formula below the data
            int totalRow = lastDataRow + 2; // leave one empty row
            sheet.Cells[totalRow, 0].PutValue("Total");
            sheet.Cells[totalRow, 1].Formula = $"SUM(B2:B{lastDataRow + 1})";

            // Add a column chart below the total row
            int chartTopRow = totalRow + 2;
            int chartIndex = sheet.Charts.Add(ChartType.Column, chartTopRow, 0, chartTopRow + 15, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Amount by Category";

            // Define the series: Y‑values from Amount column
            chart.NSeries.Add($"B2:B{lastDataRow + 1}", true);
            // Category (X) data is taken from the first column by default; explicit setting is optional.
            // chart.NSeries[0].CategoryData = $"A2:A{lastDataRow + 1}"; // Not required for basic column chart

            // Save the workbook
            string outputPath = "ChartFromXml.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
