// Title: How to automatically re‑encrypt an Excel workbook after each modification using Aspose.Cells for .NET
// AI Prompts: Create a C# wrapper that intercepts any workbook change (cell update, worksheet addition) and re‑applies Workbook.Settings.Password before saving. | Extend the SecureWorkbook class so that it calls the encryption routine automatically after each edit operation. | Show a complete example of loading an encrypted XLSX, modifying data, and saving it while preserving the original password with Aspose.Cells.
// Common Searches: Aspose.Cells .NET re‑encrypt Excel file after editing without losing password | C# automatically apply password protection when saving a modified workbook with Aspose.Cells | how to preserve workbook encryption when adding worksheets using Aspose.Cells | load encrypted XLSX, change cell value, and save with same password in C# Aspose.Cells
// Tags: auto re‑encrypt workbook after modification Aspose.Cells | Workbook.Settings.Password encryption C# | save encrypted XLSX with Aspose.Cells | modify protected Excel workbook .NET | secure Excel handling Aspose.Cells wrapper

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Required for OoxmlSaveOptions (if needed)

// The SecureWorkbook class loads or creates an XLSX workbook, stores a password, provides methods to edit cells and worksheets, and ensures the workbook is re‑encrypted by setting Workbook.Settings.Password before each Save call, enabling automatic protection after any modification.
public class SecureWorkbook
{
    private Workbook _workbook;
    private readonly string _password;

    // Load an existing workbook and apply initial encryption
    public SecureWorkbook(string filePath, string password)
    {
        try
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            // Load encrypted workbook using the provided password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            _workbook = new Workbook(filePath, loadOptions);
            _password = password;
        }
        catch (Exception ex)
        {
            throw new ApplicationException("Failed to load workbook.", ex);
        }
    }

    // Create a new workbook with encryption
    public SecureWorkbook(string password)
    {
        try
        {
            _workbook = new Workbook();
            _password = password;
        }
        catch (Exception ex)
        {
            throw new ApplicationException("Failed to create workbook.", ex);
        }
    }

    // Example modification: set a value in a cell
    public void SetCellValue(string sheetName, int row, int column, object value)
    {
        try
        {
            Worksheet sheet = _workbook.Worksheets[sheetName];
            if (sheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' does not exist.");

            sheet.Cells[row, column].PutValue(value);
        }
        catch (Exception ex)
        {
            throw new ApplicationException("Failed to set cell value.", ex);
        }
    }

    // Example modification: add a new worksheet
    public void AddWorksheet(string sheetName)
    {
        try
        {
            _workbook.Worksheets.Add(sheetName);
        }
        catch (Exception ex)
        {
            throw new ApplicationException("Failed to add worksheet.", ex);
        }
    }

    // Save the workbook to a file with encryption applied
    public void Save(string outputPath)
    {
        try
        {
            // Apply password protection before saving
            _workbook.Settings.Password = _password;

            // Save as XLSX (encryption is handled by the Settings.Password)
            _workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            throw new ApplicationException($"Failed to save workbook to '{outputPath}'.", ex);
        }
    }

    // Expose the underlying workbook for advanced operations
    public Workbook Workbook => _workbook;
}

// Usage example
class Program
{
    static void Main()
    {
        try
        {
            string password = "StrongPassword123";

            // Create a new encrypted workbook
            var sb = new SecureWorkbook(password);
            sb.AddWorksheet("Data");
            sb.SetCellValue("Data", 0, 0, "Hello");
            sb.SetCellValue("Data", 1, 0, 12345);
            sb.Save("EncryptedWorkbook.xlsx");

            // Load an existing encrypted workbook, modify, and re‑save
            var sbLoaded = new SecureWorkbook("EncryptedWorkbook.xlsx", password);
            sbLoaded.SetCellValue("Data", 2, 0, DateTime.Now);
            sbLoaded.Save("EncryptedWorkbook_Modified.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
