// Title: Convert HTML to a password‑protected XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert an HTML file to an encrypted XLSX with Aspose.Cells | how to set a password on an Excel workbook generated from HTML using Aspose.Cells .NET | Aspose.Cells example for loading HTML and saving as password‑protected Excel file | protect Excel file with a password after converting HTML in C# | save HTML as XLSX with workbook password using Aspose.Cells library
// Tags: html to xlsx conversion Aspose.Cells | apply workbook password Aspose.Cells | password protect excel file C# | load html into workbook Aspose.Cells | save workbook as xlsx with encryption Aspose.Cells

using Aspose.Cells;

// The sample loads an HTML document into an Aspose.Cells Workbook, assigns a password via the Settings.Password property, and saves the workbook as a password‑protected XLSX file.
class Program
{
    static void Main()
    {
        // Load the HTML file into a new workbook
        Workbook workbook = new Workbook("input.html");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "MySecurePassword";

        // Save the workbook as an Excel file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
