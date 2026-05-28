using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsAsyncPivotRefresh
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Paths for input and output workbooks
            string sourcePath = "PivotData.xlsx";
            string outputPath = "PivotData_Refreshed.xlsx";

            try
            {
                // Refresh pivot tables asynchronously
                await RefreshPivotTablesAsync(sourcePath, outputPath);
                Console.WriteLine("Pivot tables refreshed and workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Asynchronously refreshes all pivot tables in the workbook and saves the result.
        /// </summary>
        /// <param name="inputFile">Path to the input workbook.</param>
        /// <param name="outputFile">Path to save the refreshed workbook.</param>
        private static async Task RefreshPivotTablesAsync(string inputFile, string outputFile)
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            Workbook workbook = null;

            try
            {
                // Load the workbook on a background thread
                workbook = await Task.Run(() => new Workbook(inputFile));

                // Refresh all pivot tables using the specified options
                await Task.Run(() =>
                {
                    var refreshOption = new PivotTableRefreshOption
                    {
                        IsKeepOriginalOrder = true,
                        ReserveMissingPivotItemType = ReserveMissingPivotItemType.All
                    };

                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.RefreshPivotTables(refreshOption);
                    }
                });

                // Save the refreshed workbook on a background thread
                await Task.Run(() => workbook.Save(outputFile));
            }
            catch (Exception)
            {
                // Rethrow to allow the caller to handle/log the exception
                throw;
            }
            finally
            {
                // Dispose the workbook if it implements IDisposable
                if (workbook is IDisposable disposable)
                    disposable.Dispose();
            }
        }
    }
}