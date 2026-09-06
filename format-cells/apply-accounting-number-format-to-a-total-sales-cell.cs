// Title: Apply the built‑in Accounting number format (ID 44) to a total sales cell in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing Excel file with Aspose.Cells, sets the Accounting number format (ID 44) on cell B10, and saves the workbook. | Show how to retrieve a cell's style, change its Number property to the built‑in Accounting format, and apply the updated style using Aspose.Cells for .NET. | Demonstrate formatting a total‑sales column as accounting in a spreadsheet and exporting the result to a new file with Aspose.Cells.
// Common Searches: Aspose.Cells C# change cell B10 style to accounting and export workbook | Example of using Aspose.Cells to format a total‑sales cell as accounting | How to modify number format of a specific Excel cell with Aspose.Cells .NET | Saving Excel after updating cell style with Aspose.Cells in C#
// Tags: Aspose.Cells set accounting number format C# | apply built‑in number format ID 44 Aspose.Cells | format Excel cell as accounting .NET | save workbook after style change Aspose.Cells | total sales cell accounting style C#

using Aspose.Cells;

 // Load the existing workbook
 Workbook workbook = new Workbook("SalesReport.xlsx");

 // Access the first worksheet (or specify by name)
 Worksheet worksheet = workbook.Worksheets[0];

 // Reference the cell that contains the total sales (e.g., B10)
 Cell totalSalesCell = worksheet.Cells["B10"];

 // Retrieve the current style of the cell
 Style accountingStyle = totalSalesCell.GetStyle();

 // Apply the built‑in Accounting number format (ID 44)
 accountingStyle.Number = 44; // Accounting format

 // Set the updated style back to the cell
 totalSalesCell.SetStyle(accountingStyle);

 // Save the workbook with the applied format
 workbook.Save("SalesReport_Formatted.xlsx");
