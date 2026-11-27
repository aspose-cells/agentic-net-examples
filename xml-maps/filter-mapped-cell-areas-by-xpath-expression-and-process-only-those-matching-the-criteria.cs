// Title: Filter XML‑mapped cell areas with XPath, style matching rows and add comments using Aspose.Cells for .NET (C#)
// Common Searches: how to retrieve cell areas of an XML map query with XPath in Aspose.Cells C# | C# code to filter rows of an XML‑mapped table where a column contains a specific character | apply conditional formatting to rows returned by GetAreasOfXmlMapQuery in Aspose.Cells | add comments to cells after processing XML map rows in a .NET workbook | example of using Aspose.Cells to import XML data and style rows based on XPath criteria
// Tags: filter xml mapped cell areas using xpath Aspose.Cells | apply green style to xml mapped rows C# | insert comments into cells after xpath processing .NET | extract mapped cell ranges from xml map query | select xml mapped rows containing specific character

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace XmlMapAreaFilterDemo
{
    // The sample loads or creates an Excel workbook, adds an XML map from a schema, imports XML data, obtains cell areas that match the XPath '/Root/Item', filters rows where the 'Name' column contains the letter 'a' (case‑insensitive), applies a light‑green background style to each filtered row, adds a comment to the 'Name' cell, and saves the modified workbook as 'output.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            const string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                    <xs:element name='Root'>
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name='Item' maxOccurs='unbounded'>
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name='Name' type='xs:string'/>
                                            <xs:element name='Value' type='xs:integer'/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:element>
                </xs:schema>";

            const string xmlData = @"
                <Root>
                    <Item>
                        <Name>Alpha</Name>
                        <Value>10</Value>
                    </Item>
                    <Item>
                        <Name>Beta</Name>
                        <Value>20</Value>
                    </Item>
                    <Item>
                        <Name>Gamma</Name>
                        <Value>30</Value>
                    </Item>
                </Root>";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if present; otherwise create a new one.
                if (File.Exists(inputFile))
                {
                    workbook = new Workbook(inputFile);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Ensure at least one worksheet exists.
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "Data";

                // Add XML map if none exists.
                XmlMap xmlMap;
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                    xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                }
                else
                {
                    xmlMap = workbook.Worksheets.XmlMaps[0];
                }

                // Import XML data into the worksheet using the map.
                workbook.ImportXml(xmlData, worksheet.Name, 0, 0);

                // Query mapped areas for the desired XPath.
                string xpath = "/Root/Item";
                CellArea[] mappedAreas = worksheet.GetAreasOfXmlMapQuery(xpath, xmlMap);
                Console.WriteLine($"Total mapped areas for XPath '{xpath}': {mappedAreas.Length}");

                // Filter rows where the "Name" cell contains the letter 'a' (case‑insensitive).
                List<CellArea> areasToProcess = new List<CellArea>();
                foreach (CellArea area in mappedAreas)
                {
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        int nameCol = area.StartColumn; // "Name" is the first column of the area.
                        Cell nameCell = worksheet.Cells[row, nameCol];
                        string nameValue = nameCell.StringValue ?? string.Empty;

                        if (nameValue.IndexOf('a', StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Add the whole row as a separate CellArea.
                            areasToProcess.Add(new CellArea
                            {
                                StartRow = row,
                                EndRow = row,
                                StartColumn = area.StartColumn,
                                EndColumn = area.EndColumn
                            });
                        }
                    }
                }

                Console.WriteLine($"Areas matching filter criteria: {areasToProcess.Count}");

                // Apply style and add a comment to each filtered row.
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightGreen;
                style.Pattern = BackgroundType.Solid;

                foreach (CellArea procArea in areasToProcess)
                {
                    for (int r = procArea.StartRow; r <= procArea.EndRow; r++)
                    {
                        // Apply style to all cells in the row.
                        for (int c = procArea.StartColumn; c <= procArea.EndColumn; c++)
                        {
                            Cell cell = worksheet.Cells[r, c];
                            cell.SetStyle(style);
                        }

                        // Add a comment to the "Name" cell.
                        Cell nameCell = worksheet.Cells[r, procArea.StartColumn];
                        try
                        {
                            int commentIndex = worksheet.Comments.Add(nameCell.Name);
                            Comment comment = worksheet.Comments[commentIndex];
                            comment.Author = "Processed";
                            comment.Note = "Filtered by XPath and name contains 'a'";
                        }
                        catch (Exception commentEx)
                        {
                            Console.WriteLine($"Failed to add comment to cell {nameCell.Name}: {commentEx.Message}");
                        }
                    }
                }

                // Save the result.
                workbook.Save(outputFile);
                Console.WriteLine($"Processing complete. Output saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
