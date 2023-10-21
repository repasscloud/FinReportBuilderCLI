using OfficeOpenXml;
using FinReportBuilderCLI.Services;

namespace FinReportBuilderCLI
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Syncfusion License
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mjc2MDIzMUAzMjMzMmUzMDJlMzBFdGdaRVVWL1duUyt1TERYK3kydjhNOXl3ck42Q3Y1eWRMMjR3UlJJbnRFPQ==");

            // EPP License
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            string clientName = "zaimay pty ltd";
            string abn = "82144820897";
            string acn = "144820897";
            double retainedEarningsLastFiscalYear = 6147.00;
            double dividendPaidLastFiscalYear = 0.00;
            double dividendPaidThisFiscalYear = 0.00;

            // Get the current directory where the executable is located
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string excelFilePath = "Book1.xlsx";
            FileInfo fileInfo = new FileInfo(Path.Combine(currentDirectory, excelFilePath));

            Console.WriteLine(Path.Combine(currentDirectory, excelFilePath));
            ExcelPackage excel = new ExcelPackage(fileInfo);

            // int numColumns = 0;
            // int numRows = 0;
            // List<string> columnNameText = new();

            // Read excel file
            // using (ExcelPackage package = new ExcelPackage(fileInfo))
            // {
            //     // Get  the first worksheet
            //     ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //     // determine the number of columns in worksheet
            //     numColumns = worksheet.Dimension.End.Column;

            //     // determine the number of rows in worksheet
            //     numRows = worksheet.Dimension.End.Column;

            //     // get the column names
            //     string[] columnNames;
            //     columnNames = new string[numColumns];
            //     for (int i = 1; i <= numColumns; i++)
            //     {
            //         columnNames[i - 1] = worksheet.Cells[1, i].Text;
            //         columnNameText.Add(worksheet.Cells[1, i].Text);
            //     }

            //     //for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            //     //{
            //     //    for (int col = 1; col <= numColumns; col++)
            //     //    {
            //     //        string columnName = columnNames[col - 1];
            //     //        string data = worksheet.Cells[row, col].Text;
            //     //        Console.WriteLine($"{columnName}: {data}");
            //     //    }
            //     //}
            // }

            FinancialReportService reportService = new FinancialReportService();

            // Create the financial report
            var documentStream = reportService.CreateFinancialReportForYearEnded(clientName, abn, acn, 
                retainedEarningsLastFiscalYear, dividendPaidLastFiscalYear, dividendPaidThisFiscalYear, fileInfo);

            // Specify the path where you want to save the document
            string filePath = Path.Combine(currentDirectory, "FinancialReport.docx"); // Replace with your desired file path

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                documentStream.WriteTo(fileStream);
            }

            // Close the streams
            documentStream.Close();

            Console.WriteLine("Financial report saved to: " + filePath);
        }
    }
}

