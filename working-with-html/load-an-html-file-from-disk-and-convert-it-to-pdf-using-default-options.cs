using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Desired output PDF file path
        string pdfFilePath = "output.pdf";

        // Convert the HTML file to PDF using default load and save options
        // This utilizes the ConversionUtility.Convert method which handles loading and saving internally.
        ConversionUtility.Convert(htmlFilePath, pdfFilePath);

        Console.WriteLine($"HTML file '{htmlFilePath}' has been successfully converted to PDF '{pdfFilePath}'.");
    }
}