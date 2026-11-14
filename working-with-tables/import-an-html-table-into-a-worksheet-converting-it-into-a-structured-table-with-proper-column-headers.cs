// Title: Import an HTML table into an Excel worksheet and convert it to a styled ListObject using Aspose.Cells for .NET (C#)
// Common Searches: how to load an HTML string into an Aspose.Cells workbook in C# | convert HTML table to Excel ListObject with headers using Aspose.Cells | add a structured table to a worksheet after importing HTML with Aspose.Cells | apply a built‑in table style to an imported HTML table in Aspose.Cells C# | save imported HTML table as .xlsx using Aspose.Cells .NET
// Tags: HtmlLoadOptions for importing HTML content | initialize ListObject from worksheet used range | set TableStyleMedium9 on Aspose.Cells table | transform HTML table into Aspose.Cells ListObject | save workbook as XLSX file with Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an HTML string containing a table into an Aspose.Cells Workbook via HtmlLoadOptions, determines the used cell range on the first worksheet, creates a ListObject (structured Excel table) covering that range with header rows, applies the built‑in TableStyleMedium9, and saves the workbook as ImportedTable.xlsx, handling directory creation and exceptions.
class HtmlTableImport
{
    static void Main()
    {
        try
        {
            // Sample HTML containing a table.
            string html = @"
            <html>
                <body>
                    <table border='1'>
                        <tr>
                            <th>Product</th>
                            <th>Quantity</th>
                            <th>Price</th>
                        </tr>
                        <tr>
                            <td>Apple</td>
                            <td>10</td>
                            <td>0.5</td>
                        </tr>
                        <tr>
                            <td>Banana</td>
                            <td>20</td>
                            <td>0.3</td>
                        </tr>
                    </table>
                </body>
            </html>";

            // Load HTML into a workbook using HtmlLoadOptions.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                Workbook workbook = new Workbook(ms, loadOptions);

                // Get the first worksheet where the HTML table was imported.
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range (including headers).
                int totalRows = sheet.Cells.MaxDataRow + 1;      // MaxDataRow is zero‑based.
                int totalCols = sheet.Cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based.

                // Add a structured table (ListObject) based on the used range.
                // Use overload that accepts row/column counts because older API versions may not support Range overload.
                int listObjectIndex = sheet.ListObjects.Add(0, 0, totalRows, totalCols, true);
                ListObject listObject = sheet.ListObjects[listObjectIndex];

                // Apply a predefined table style.
                listObject.TableStyleType = TableStyleType.TableStyleMedium9;

                // Prepare output path and ensure directory exists.
                string outputPath = "ImportedTable.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
