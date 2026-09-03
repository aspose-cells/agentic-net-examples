// Title: Highlight cells with a specific StyleID in an Excel workbook using XPath on SpreadsheetML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, converts it to SpreadsheetML, runs an XPath query to find cells having a given StyleID, and applies a yellow background style to those cells. | Show how to parse the XPath result from the SpreadsheetML XML, build cell addresses, and set a custom style on each matching cell using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# filter cells by XPath and change background color | How to select cells with a specific StyleID using XPath in SpreadsheetML with Aspose.Cells | Apply style to cells based on XML mapping in Aspose.Cells .NET
// Tags: XPath cell selection Aspose.Cells .NET | apply background style to cells Aspose.Cells | SpreadsheetML conversion Aspose.Cells C# | highlight cells by StyleID Aspose.Cells | C# extract cell addresses from XML Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml;
using Aspose.Cells;

// The example loads or creates an Excel workbook, converts it to SpreadsheetML in memory, uses an XPath expression to locate cells with a particular StyleID, builds a set of zero‑based cell addresses, creates a yellow background style, applies that style to each matching cell, and saves the modified workbook as a new XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook; create a new one if the input file does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Convert workbook to SpreadsheetML (XML) in memory
            XmlDocument xmlDoc = new XmlDocument();
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xml);
                ms.Position = 0;
                xmlDoc.Load(ms);
            }

            // XPath that selects cells with a specific style ID (example)
            const string xpath = "//Cell[@ss:StyleID='1']";

            // Register SpreadsheetML namespace
            XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsMgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");

            // Execute XPath query
            XmlNodeList matchingNodes = xmlDoc.SelectNodes(xpath, nsMgr);

            // Build a set of cell addresses that match the XPath criteria.
            // Address format: "R{rowIndex}C{colIndex}" (zero‑based)
            HashSet<string> matchingAddresses = new HashSet<string>();
            if (matchingNodes != null)
            {
                foreach (XmlNode cellNode in matchingNodes)
                {
                    if (cellNode?.ParentNode == null) continue;

                    XmlNode rowNode = cellNode.ParentNode;

                    int rowIndex = -1;
                    int colIndex = -1;

                    // Row index (1‑based) from <Row ss:Index="...">
                    if (rowNode.Attributes != null &&
                        rowNode.Attributes["ss:Index"] != null &&
                        int.TryParse(rowNode.Attributes["ss:Index"]?.Value, out int rIdx))
                    {
                        rowIndex = rIdx - 1;
                    }

                    // Column index (1‑based) from <Cell ss:Index="...">
                    if (cellNode.Attributes != null &&
                        cellNode.Attributes["ss:Index"] != null &&
                        int.TryParse(cellNode.Attributes["ss:Index"]?.Value, out int cIdx))
                    {
                        colIndex = cIdx - 1;
                    }

                    // If both indices are available, add the address
                    if (rowIndex >= 0 && colIndex >= 0)
                    {
                        matchingAddresses.Add($"R{rowIndex}C{colIndex}");
                    }
                }
            }

            // Prepare a style to apply (yellow background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to each matching cell
            foreach (string address in matchingAddresses)
            {
                // Expected format: R{row}C{col}
                int rPos = address.IndexOf('R');
                int cPos = address.IndexOf('C');
                if (rPos == -1 || cPos == -1) continue;

                if (int.TryParse(address.Substring(rPos + 1, cPos - rPos - 1), out int row) &&
                    int.TryParse(address.Substring(cPos + 1), out int col))
                {
                    // Ensure the cell exists (Aspose.Cells creates it on demand)
                    Cell cell = sheet.Cells[row, col];
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
