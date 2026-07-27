using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ContentTypePropertiesReport
{
    static void Main(string[] args)
    {
        // Folder containing the workbooks to scan.
        string folderPath = @"C:\Workbooks";

        // Ensure the folder exists.
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Output CSV file.
        string reportPath = Path.Combine(folderPath, "ContentTypePropertiesReport.csv");

        // Prepare CSV header.
        var sb = new StringBuilder();
        sb.AppendLine("Workbook,PropertyName,IsNillable");

        // Enumerate all Excel files in the folder (including subfolders).
        foreach (string file in Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories))
        {
            // Ensure the file actually exists.
            if (!File.Exists(file))
                continue;

            try
            {
                // Load the workbook. If the file is password‑protected,
                // Aspose.Cells throws a CellsException.
                var loadOptions = new LoadOptions();
                using (Workbook wb = new Workbook(file, loadOptions))
                {
                    bool hasMissingFlag = false;

                    // Iterate through all content type properties.
                    for (int i = 0; i < wb.ContentTypeProperties.Count; i++)
                    {
                        var prop = wb.ContentTypeProperties[i]; // Use var to avoid explicit type issues.

                        // Record properties where IsNillable is false.
                        if (!prop.IsNillable)
                        {
                            hasMissingFlag = true;
                            sb.AppendLine($"{Path.GetFileName(file)},{prop.Name},{prop.IsNillable}");
                        }
                    }

                    // Handle workbooks with no ContentTypeProperties.
                    if (wb.ContentTypeProperties.Count == 0)
                    {
                        sb.AppendLine($"{Path.GetFileName(file)},<No ContentTypeProperties>,");
                    }
                    else if (!hasMissingFlag)
                    {
                        // All properties are nillable.
                        sb.AppendLine($"{Path.GetFileName(file)},<All IsNillable>,True");
                    }
                }
            }
            catch (CellsException ex)
            {
                // Password‑protected files throw a CellsException whose message contains
                // the word "Password". Treat them as such.
                if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sb.AppendLine($"{Path.GetFileName(file)},<Password Protected>,");
                }
                else
                {
                    sb.AppendLine($"{Path.GetFileName(file)},<CellsException: {ex.Message}>,");
                }
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors and continue processing.
                sb.AppendLine($"{Path.GetFileName(file)},<Error: {ex.Message}>,");
            }
        }

        // Write the CSV report to disk.
        try
        {
            File.WriteAllText(reportPath, sb.ToString());
            Console.WriteLine($"Report generated: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }
    }
}