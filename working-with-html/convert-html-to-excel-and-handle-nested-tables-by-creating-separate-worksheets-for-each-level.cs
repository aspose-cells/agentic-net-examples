// Title: Convert an HTML file with nested tables into separate worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML file, locates every <table> element regardless of nesting, and creates a dedicated worksheet for each table with Aspose.Cells. | Write a C# routine that names each worksheet using its nesting depth and a sequential counter while importing the table data into the sheet. | Provide C# logic to ignore tables that contain no rows and output the processing hierarchy of tables to the console.
// Common Searches: asp.net convert html file with nested tables to excel workbook using aspose.cells | c# parse html tables recursively and write each to its own worksheet | how to export each html table level to a separate worksheet in .net | aspose.cells generate multiple worksheets from html based on table hierarchy | c# read html and create excel sheets for every table element
// Tags: Aspose.Cells HTML to Excel worksheet conversion | nested HTML table extraction with XDocument | recursive table processing for Excel export | C# create worksheet per HTML table level | Aspose.Cells import two‑dimensional array from HTML

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;

namespace HtmlToExcelNestedTables
{
    // The example reads an HTML file, wraps it in a root element, and parses it with XDocument. It identifies top‑level tables and recursively processes each table and any nested tables. For every table a uniquely named worksheet is added to an Aspose.Cells workbook, the cell text is extracted, and the data is imported via a two‑dimensional array. After all tables are handled, the workbook is saved as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Verify that the input file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: File '{htmlPath}' not found.");
                    return;
                }

                // Load the HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Wrap the HTML with a root element to make it well‑formed XML
                string wrappedHtml = $"<root>{htmlContent}</root>";
                XDocument xDoc = XDocument.Parse(wrappedHtml);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Counter for unique worksheet names
                int sheetCounter = 1;

                // Process top‑level tables (those directly under the body)
                var topTables = xDoc.Descendants("body")
                                    .Descendants("table")
                                    .Where(t => t.Ancestors("table").Count() == 0); // only top‑level

                foreach (var tableElem in topTables)
                {
                    ProcessTable(tableElem, workbook, "Table", 1, ref sheetCounter);
                }

                // Save the workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved as 'output.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="tableElem">The XML element representing the HTML table.</param>
        /// <param name="workbook">The Aspose.Cells workbook.</param>
        /// <param name="baseName">Base name for worksheets.</param>
        /// <param name="level">Current nesting level (1 = top level).</param>
        /// <param name="sheetCounter">Reference counter to ensure unique sheet names.</param>
        private static void ProcessTable(XElement tableElem, Workbook workbook, string baseName, int level, ref int sheetCounter)
        {
            try
            {
                // Create a new worksheet for this table
                string sheetName = $"{baseName}_L{level}_{sheetCounter}";
                Worksheet sheet = workbook.Worksheets.Add(sheetName);
                sheetCounter++;

                // Extract rows (tr) from the table, handling optional tbody/thead/tfoot wrappers
                var rowElements = tableElem.Elements()
                                           .Where(e => e.Name == "tr" ||
                                                       e.Name == "tbody" ||
                                                       e.Name == "thead" ||
                                                       e.Name == "tfoot")
                                           .SelectMany(e => e.Name == "tr" ? new[] { e } : e.Elements("tr"))
                                           .ToList();

                // Collect cell values
                List<List<string>> rowsData = new List<List<string>>();
                foreach (var row in rowElements)
                {
                    List<string> cellValues = new List<string>();
                    var cells = row.Elements().Where(c => c.Name == "th" || c.Name == "td");
                    foreach (var cell in cells)
                    {
                        string cellText = (cell.Value ?? string.Empty).Trim()
                                             .Replace("\r", " ")
                                             .Replace("\n", " ");
                        cellValues.Add(cellText);
                    }
                    rowsData.Add(cellValues);
                }

                // Determine the maximum column count
                int maxCols = rowsData.Any() ? rowsData.Max(r => r.Count) : 0;

                // Build a rectangular array for import
                object[,] dataArray = new object[rowsData.Count, maxCols];
                for (int i = 0; i < rowsData.Count; i++)
                {
                    for (int j = 0; j < rowsData[i].Count; j++)
                    {
                        dataArray[i, j] = rowsData[i][j];
                    }
                }

                // Import the data into the worksheet starting at cell A1
                if (rowsData.Count > 0 && maxCols > 0)
                {
                    sheet.Cells.ImportTwoDimensionArray(dataArray, 0, 0);
                }

                // Find nested tables within the current table (excluding the current table itself)
                var nestedTables = tableElem.Descendants("table")
                                            .Where(t => t != tableElem);
                foreach (var nestedTable in nestedTables)
                {
                    // Recursively process the nested table at the next level
                    ProcessTable(nestedTable, workbook, baseName, level + 1, ref sheetCounter);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing table at level {level}: {ex.Message}");
            }
        }
    }
}
