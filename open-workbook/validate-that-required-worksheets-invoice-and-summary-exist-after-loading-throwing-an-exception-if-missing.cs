// Title: How to validate that "Invoice" and "Summary" worksheets exist in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if worksheet named Invoice exists in workbook | C# validate required sheets Invoice and Summary before processing Excel file | Throw exception when mandatory worksheet is missing using Aspose.Cells .NET | How to ensure specific worksheets are present in an Excel file with Aspose.Cells | Load Excel workbook and verify required sheet names with Aspose.Cells API
// Tags: Aspose.Cells worksheet existence validation | C# required worksheet check Excel | Aspose.Cells missing sheet exception handling | validate mandatory sheets Aspose.Cells .NET | load workbook verify sheet names Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook with Aspose.Cells in C#, checking for the presence of the "Invoice" and "Summary" worksheets, and throwing a detailed exception if either is missing. It also includes file‑not‑found handling and wraps other errors, confirming successful validation with a console message.
public class WorkbookValidator
{
    /// <param name="filePath">Path to the Excel file to validate.</param>
    public void ValidateRequiredSheets(string filePath)
    {
        // Ensure the file exists before attempting to load.
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"The file \"{filePath}\" was not found.");
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Attempt to retrieve the required worksheets by name.
            Worksheet invoiceSheet = workbook.Worksheets["Invoice"];
            Worksheet summarySheet = workbook.Worksheets["Summary"];

            // If either worksheet is not found, throw an informative exception.
            if (invoiceSheet == null || summarySheet == null)
            {
                string missingSheets = "";
                if (invoiceSheet == null) missingSheets += "\"Invoice\" ";
                if (summarySheet == null) missingSheets += "\"Summary\" ";
                throw new Exception($"The workbook is missing required worksheet(s): {missingSheets.Trim()}.");
            }

            // Both required worksheets are present.
            Console.WriteLine("Validation succeeded: Required worksheets are present.");
        }
        catch (Exception ex) when (!(ex is FileNotFoundException))
        {
            // Wrap any loading or processing errors.
            throw new Exception($"An error occurred while validating the workbook: {ex.Message}", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage: provide the Excel file path as the first argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as a command-line argument.");
            return;
        }

        string filePath = args[0];
        WorkbookValidator validator = new WorkbookValidator();

        try
        {
            validator.ValidateRequiredSheets(filePath);
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Validation error: {ex.Message}");
        }
    }
}
