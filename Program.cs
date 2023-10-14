using FinReportBuilderCLI.Services;

namespace FinReportBuilderCLI
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Syncfusion License
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");

            string clientName = "zaimay pty ltd";
            string abn = "82144820897";
            string acn = "144820897";

            FinancialReportService reportService = new FinancialReportService();

            // Create the financial report
            var documentStream = reportService.CreateFinancialReportForYearEnded(clientName, abn, acn);

            // Specify the path where you want to save the document
            string filePath = "C:\\Users\\Danijel.Wynyard\\tmp\\FinancialReport.docx"; // Replace with your desired file path


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