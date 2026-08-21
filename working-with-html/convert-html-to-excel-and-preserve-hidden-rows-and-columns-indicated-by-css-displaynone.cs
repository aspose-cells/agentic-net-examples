// Title: C# – Convert HTML (display:none) to Excel with hidden rows/columns using Aspose.Cells
// Description: This C# example demonstrates how Aspose.Cells for .NET loads an HTML file, interprets CSS display:none as hidden rows and columns, and saves the workbook to XLSX while keeping those rows and columns hidden.
// Keywords: Aspose.Cells | C# | HTML to Excel | preserve hidden rows | preserve hidden columns | CSS display:none | load HTML workbook | save as XLSX | Aspose.Cells for .NET | convert HTML tables
// Common Searches: Aspose.Cells keep hidden rows when converting HTML to Excel | C# convert HTML with display:none to XLSX preserving hidden columns | load HTML file into Aspose.Cells workbook hidden rows | retain CSS hidden elements in Excel using Aspose.Cells
// Developer Intent: Convert an HTML document that hides rows or columns with CSS display:none into an Excel file while maintaining the hidden state.
// Use Cases: Export web‑based reports that contain hidden sections to Excel without exposing the hidden data. | Migrate legacy HTML tables that rely on CSS visibility rules into Excel worksheets for further analysis. | Automate batch processing of HTML templates, generating Excel files that preserve hidden rows and columns.
// AI Prompts: Show code to list all hidden rows and columns after loading the HTML workbook with Aspose.Cells. | Demonstrate how to keep cell formatting (e.g., background colors) while converting HTML to Excel using Aspose.Cells. | Provide a script that processes multiple HTML files in a directory, converting each to XLSX and preserving hidden rows/columns.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    // This C# example demonstrates how Aspose.Cells for .NET loads an HTML file, interprets CSS display:none as hidden rows and columns, and saves the workbook to XLSX while keeping those rows and columns hidden.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains CSS display:none for hidden rows/columns
            string htmlFilePath = "input.html";

            // Load the HTML file into a Workbook.
            // Aspose.Cells parses the HTML and interprets CSS display:none as hidden rows/columns.
            Workbook workbook = new Workbook(htmlFilePath);

            // Optional: verify that hidden rows/columns are recognized (for debugging)
            // Console.WriteLine("Hidden rows count: " + workbook.Worksheets[0].Cells.GetHiddenRows().Count);
            // Console.WriteLine("Hidden columns count: " + workbook.Worksheets[0].Cells.GetHiddenColumns().Count);

            // Save the workbook as an Excel file, preserving the hidden rows and columns.
            string excelFilePath = "output.xlsx";
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel with hidden rows/columns preserved.");
        }
    }
}
