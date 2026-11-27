// Title: Apply a custom highlight style to cells mapped from the /Customer/Name XML element using Aspose.Cells for .NET (C#)
// Common Searches: how to highlight customer name cells when importing XML with Aspose.Cells C# | Aspose.Cells apply background color and bold font to cells mapped from an XML element | C# code to style specific Excel cells based on XML mapping using Aspose.Cells | set custom cell style for /Customer/Name nodes in an Excel workbook with Aspose.Cells | auto fit columns after styling XML‑mapped cells in Aspose.Cells .NET
// Tags: custom style for XML‑mapped cells Aspose.Cells | highlighting cells from /Customer/Name element | auto‑fit worksheet columns Aspose.Cells | Excel workbook cell formatting C# Aspose.Cells | XML to Excel data import with cell styling Aspose.Cells

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using System.Drawing;

namespace ApplyCustomStyleToCustomerNames
{
    // The example loads a customers.xml file, creates a new workbook, defines a light‑yellow background with bold dark‑blue font style, writes a header row, iterates through each <Customer> element, writes the ID and Name into columns A and B, applies the custom highlight style to each Name cell, auto‑fits the columns, and saves the workbook as CustomerNames.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input XML and output Excel file
            const string xmlFilePath = "customers.xml";
            const string outputExcelPath = "CustomerNames.xlsx";

            // Verify that the XML file exists
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"Error: XML file not found at path '{xmlFilePath}'.");
                return;
            }

            try
            {
                // Load the XML document
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a custom style for highlighting customer names
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.LightYellow;
                highlightStyle.Pattern = BackgroundType.Solid;
                highlightStyle.Font.IsBold = true;
                highlightStyle.Font.Color = Color.DarkBlue;

                // Optional: write a header
                cells["A1"].PutValue("Customer ID");
                cells["B1"].PutValue("Customer Name");
                // Apply a simple header style
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                cells["A1"].SetStyle(headerStyle);
                cells["B1"].SetStyle(headerStyle);

                // Iterate over each /Customer element
                XmlNodeList customerNodes = xmlDoc.SelectNodes("/Customers/Customer");
                if (customerNodes == null)
                {
                    Console.WriteLine("No <Customer> elements found in the XML.");
                    return;
                }

                int currentRow = 2; // Start after header
                foreach (XmlNode customerNode in customerNodes)
                {
                    // Extract ID (if present) and Name
                    string id = customerNode.SelectSingleNode("ID")?.InnerText ?? string.Empty;
                    string name = customerNode.SelectSingleNode("Name")?.InnerText ?? string.Empty;

                    // Write ID and Name to the worksheet
                    cells[$"A{currentRow}"].PutValue(id);
                    Cell nameCell = cells[$"B{currentRow}"];
                    nameCell.PutValue(name);

                    // Apply the custom highlight style to the Name cell
                    nameCell.SetStyle(highlightStyle);

                    currentRow++;
                }

                // Auto-fit columns for better appearance
                worksheet.AutoFitColumn(0);
                worksheet.AutoFitColumn(1);

                // Save the workbook
                workbook.Save(outputExcelPath);
                Console.WriteLine($"Workbook saved successfully to '{outputExcelPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
