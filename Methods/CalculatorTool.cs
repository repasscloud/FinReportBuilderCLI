using Syncfusion.DocIO.DLS;

namespace FinReportBuilderCLI.Methods
{
	public class CalculatorTool
	{
		public static double CalculateTaxRate(int thisFiscalYear, double thisFiscalYearPBTI)
        {
            if (thisFiscalYear >= 2017 && thisFiscalYear <= 2018)
            {
                return thisFiscalYearPBTI < 25000000 ? 0.275 : 0.3;
            }
            else if (thisFiscalYear >= 2019 && thisFiscalYear <= 2020)
            {
                return thisFiscalYearPBTI < 50000000 ? 0.275 : 0.3;
            }
            else if (thisFiscalYear >= 2021 && thisFiscalYear <= 2022)
            {
                return thisFiscalYearPBTI < 50000000 ? 0.26 : 0.3;
            }
            else if (thisFiscalYear >= 2023)
            {
                return thisFiscalYearPBTI < 50000000 ? 0.25 : 0.3;
            }
            else
            {
                // Handle any other cases as needed
                return 0.0; // Default value if the year is not within the specified ranges
            }
        }
	}
}

