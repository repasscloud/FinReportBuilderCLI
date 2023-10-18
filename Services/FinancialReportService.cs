using System.Globalization;
using FinReportBuilderCLI.Methods;
using OfficeOpenXml;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace FinReportBuilderCLI.Services
{
    public class FinancialReportService
    {
        public MemoryStream CreateFinancialReportForYearEnded(
            string clientName,
            string? abn,
            string? acn,
            double retainedEarningsLastFiscalYear,
            double dividendPaidLastFiscalYear,
            double dividendPaidThisFiscalYear,
            FileInfo fileInfo)
        {
            // Create new Word Document
            using (WordDocument wordDocument = new WordDocument())
            {
                #region Section01
                // Section01 - Title Page
                IWSection section01 = wordDocument.AddSection();

                // Section01 - Page Setup
                section01.PageSetup.Orientation = PageOrientation.Portrait;
                section01.PageSetup.Margins.All = 36;

                // Section01 - Paragraph Style 01 (Title)
                IWParagraphStyle secn01Style01 = wordDocument.AddParagraphStyle("Section01Style01");
                secn01Style01.ParagraphFormat.BackColor = Color.White;
                secn01Style01.ParagraphFormat.AfterSpacing = 18f;
                secn01Style01.ParagraphFormat.BeforeSpacing = 18f;
                secn01Style01.ParagraphFormat.LineSpacing = 16f;
                secn01Style01.CharacterFormat.FontName = "Times New Roman";
                secn01Style01.CharacterFormat.FontSize = 16f;
                secn01Style01.CharacterFormat.Bold = true;

                // Section01 - Paragraph Style 02 (Pre-Footer)
                IWParagraphStyle secn01Style02 = wordDocument.AddParagraphStyle("Section01Style02");
                secn01Style02.ParagraphFormat.BackColor = Color.White;
                secn01Style02.ParagraphFormat.AfterSpacing = 14f;
                secn01Style02.ParagraphFormat.BeforeSpacing = 14f;
                secn01Style02.ParagraphFormat.LineSpacing = 14f;
                secn01Style02.CharacterFormat.FontName = "Times New Roman";
                secn01Style02.CharacterFormat.FontSize = 12f;
                secn01Style02.CharacterFormat.Bold = true;

                // Section01 - Title Paragraph
                IWParagraph paragraph01 = section01.AddParagraph();
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendText($"{clientName.ToUpperInvariant()}");
                paragraph01.AppendBreak(BreakType.LineBreak);

                // Section01 - ABN/ACN
                string companyAbnAcn;
                if (!string.IsNullOrEmpty(abn) && !string.IsNullOrEmpty(acn))
                {
                    companyAbnAcn = $"ABN {long.Parse(abn!):00 000 000 000} / ACN {long.Parse(acn!):000 000 000}";
                }
                else if (!string.IsNullOrEmpty(abn))
                {
                    companyAbnAcn = $"ABN {long.Parse(abn!):00 000 000 000}";
                }
                else if (!string.IsNullOrEmpty(acn))
                {
                    companyAbnAcn = $"ACN {long.Parse(acn!):000 000 000}";
                }
                else
                {
                    companyAbnAcn = string.Empty; // Both are empty
                }
                paragraph01.AppendText(companyAbnAcn);

                // Section01 - Report Title
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendText("financial report".ToUpperInvariant());
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendText("for the year ended".ToUpperInvariant());
                paragraph01.AppendBreak(BreakType.LineBreak);
                paragraph01.AppendText("30 june 2021".ToUpperInvariant());
                paragraph01.ApplyStyle("Section01Style01");
                paragraph01.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Section01 - Pre Footer Paragraph
                IWParagraph paragraph02 = section01.AddParagraph();
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendText("Liability limited by a scheme approved under");
                paragraph02.AppendBreak(BreakType.LineBreak);
                paragraph02.AppendText("Professional Standards Legislation");
                paragraph02.ApplyStyle("Section01Style02");
                paragraph02.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                #endregion Section01

                #region Section02
                // Section02 - Table of Contents
                IWSection section02 = wordDocument.AddSection();

                // Section02 - Page Setup
                section02.PageSetup.Orientation = PageOrientation.Portrait;
                section02.PageSetup.Margins.All = 36;

                // Section01 - Paragraph Style 01 (Title)
                IWParagraphStyle secn02Style01 = wordDocument.AddParagraphStyle("Section02Style01");
                secn02Style01.ParagraphFormat.BackColor = Color.White;
                secn02Style01.ParagraphFormat.AfterSpacing = 18f;
                secn02Style01.ParagraphFormat.BeforeSpacing = 18f;
                secn02Style01.ParagraphFormat.LineSpacing = 16f;
                secn01Style02.CharacterFormat.FontName = "Times New Roman";
                secn02Style01.CharacterFormat.FontSize = 16f;
                secn02Style01.CharacterFormat.Bold = true;

                // Section02 - Title Paragraph
                IWParagraph paragraph03 = section02.AddParagraph();
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendText($"{clientName.ToUpperInvariant()}");
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendText(companyAbnAcn);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendBreak(BreakType.LineBreak);
                paragraph03.AppendText("CONTENTS");
                paragraph03.ApplyStyle("Section02Style01");
                paragraph03.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Section02 - Table of Contents
                IWTable tocTable = section02.AddTable();
                tocTable.TableFormat.Borders.BorderType = BorderStyle.None;
                tocTable.ResetCells(6, 2);

                tocTable[0, 0].AddParagraph().AppendText("Income Statement\n");
                IWParagraph tocIncomeStatement = tocTable[0, 1].AddParagraph();
                tocIncomeStatement.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocIncomeStatement.AppendText("3");

                tocTable[1, 0].AddParagraph().AppendText("Balance Sheet\n");
                IWParagraph tocBalanceSheet = tocTable[1, 1].AddParagraph();
                tocBalanceSheet.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocBalanceSheet.AppendText("4");

                tocTable[2, 0].AddParagraph().AppendText("Notes to the Financial Statements\n");
                IWParagraph tocNotesToTheFinancialStatements = tocTable[2, 1].AddParagraph();
                tocNotesToTheFinancialStatements.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocNotesToTheFinancialStatements.AppendText("5");

                tocTable[3, 0].AddParagraph().AppendText("Director's Declaration\n");
                IWParagraph tocDirectorsDeclaration = tocTable[3, 1].AddParagraph();
                tocDirectorsDeclaration.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocDirectorsDeclaration.AppendText("8");

                tocTable[4, 0].AddParagraph().AppendText("Compilation Report\n");
                IWParagraph tocCompilationReport = tocTable[4, 1].AddParagraph();
                tocCompilationReport.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocCompilationReport.AppendText("9");

                tocTable[5, 0].AddParagraph().AppendText("Detailed Profit and Loss Statement\n");
                IWParagraph tocDetailedProfitAndLossStatement = tocTable[5, 1].AddParagraph();
                tocDetailedProfitAndLossStatement.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                tocDetailedProfitAndLossStatement.AppendText("11");
                #endregion Section02

                #region Section03
                // Section03 - Title Paragraph
                IWSection section03 = wordDocument.AddSection();

                // Section03 - Page Setup
                section03.PageSetup.Orientation = PageOrientation.Portrait;
                section03.PageSetup.Margins.All = 36;

                // Section03 - Paragraph Style 03 (Title)
                IWParagraphStyle secn03Style01 = wordDocument.AddParagraphStyle("Section03Style01");
                secn03Style01.ParagraphFormat.BackColor = Color.White;
                secn03Style01.ParagraphFormat.AfterSpacing = 16f;
                secn03Style01.ParagraphFormat.BeforeSpacing = 16f;
                secn03Style01.ParagraphFormat.LineSpacing = 14f;
                secn03Style01.CharacterFormat.FontName = "Times New Roman";
                secn03Style01.CharacterFormat.FontSize = 14f;
                secn03Style01.CharacterFormat.Bold = true;

                // Section03 - Used by the "footer" of the page
                IWParagraphStyle page03Style02 = wordDocument.AddParagraphStyle("Page03Style02");
                page03Style02.ParagraphFormat.BackColor = Color.White;
                page03Style02.ParagraphFormat.AfterSpacing = 12f;
                page03Style02.ParagraphFormat.BeforeSpacing = 12f;
                page03Style02.ParagraphFormat.LineSpacing = 10f;
                page03Style02.CharacterFormat.FontName = "Times New Roman";
                page03Style02.CharacterFormat.FontSize = 10f;
                page03Style02.CharacterFormat.Bold = false;

                // Section03 - Title Paragraph
                IWParagraph paragraph04 = section03.AddParagraph();
                paragraph04.AppendText($"{clientName.ToUpperInvariant()}");
                paragraph04.AppendBreak(BreakType.LineBreak);
                paragraph04.AppendText(companyAbnAcn);
                paragraph04.AppendBreak(BreakType.LineBreak);
                paragraph04.AppendBreak(BreakType.LineBreak);
                paragraph04.AppendText("NOTES TO THE FINANCIAL STATEMENTS");
                paragraph04.AppendBreak(BreakType.LineBreak);
                paragraph04.AppendText("FOR THE YEAR ENDED 30 JUNE 2020");
                paragraph04.AppendBreak(BreakType.LineBreak);

                // Section03 - HR
                paragraph04.AppendText("__________________________________________________________________________");
                paragraph04.ApplyStyle("Section03Style01");
                paragraph04.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Page03 - Read Excel File
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet page03IncomeWorksheet = package.Workbook.Worksheets[0];
                ExcelWorksheet page03ExpenditureWorksheet = package.Workbook.Worksheets[1];

                // Page03 - Table of Contents
                IWTable page3Table = section03.AddTable();
                page3Table.TableFormat.Borders.BorderType = BorderStyle.None;
                page3Table.TableFormat.HorizontalAlignment = RowAlignment.Center;
                int page3TableTotalRowCount = 0;
                // add first row into table
                WTableRow row = page3Table.AddRow();
                page3TableTotalRowCount++;
                int page3TableCell1Width = 270;
                int page3TableCell2Width = 70;
                int page3TableCell3_4Width = 90;
                
                // add cells to first row (heading row)
                WTableCell cell = row.AddCell();
                cell.Width = page3TableCell1Width;

                cell = row.AddCell();
                cell.AddParagraph().AppendText("NOTE\n").CharacterFormat.Bold = true;
                cell.Width = page3TableCell2Width;
                page3Table.Rows[0].Cells[1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                int startCell = 2;
                for (int i = 3; i <= page03IncomeWorksheet.Dimension.End.Column; i++)
                {
                    cell = row.AddCell();
                    cell.Width = page3TableCell3_4Width;
                    cell.AddParagraph().AppendText($"{page03IncomeWorksheet.Cells[1, i].Text}\n$").CharacterFormat.Bold = true;
                    page3Table.Rows[0].Cells[startCell].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                    startCell++;
                }

                // add INCOME row to table
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.AddParagraph().AppendText("INCOME\n").CharacterFormat.Bold = true;
                cell.Width = page3TableCell1Width;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;

                // add INCOME rows to table
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                for (int i = 1; i <= page03ExpenditureWorksheet.Dimension.End.Row; i++)
                {
                    switch(i)
                    {
                        case 1:
                            cell = row.AddCell();
                            cell.Width = page3TableCell1Width;
                            cell.AddParagraph().AppendText($"{page03IncomeWorksheet.Cells[2, i].Text}");
                            page3Table.Rows[2].Cells[i - 1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                            break;
                        case 2:
                            cell = row.AddCell();
                            cell.Width = page3TableCell2Width;
                            cell.AddParagraph().AppendText($"{page03IncomeWorksheet.Cells[2, i].Text}");
                            page3Table.Rows[2].Cells[i - 1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                            break;
                        case 3:
                            cell = row.AddCell();
                            cell.Width = page3TableCell3_4Width;
                            cell.AddParagraph().AppendText($"{page03IncomeWorksheet.Cells[2, i].Text}");
                            page3Table.Rows[2].Cells[i - 1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            break;
                        case 4:
                            cell = row.AddCell();
                            cell.Width = page3TableCell3_4Width;
                            cell.AddParagraph().AppendText($"{page03IncomeWorksheet.Cells[2, i].Text}");
                            page3Table.Rows[2].Cells[i - 1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            break;
                        default:
                            break;
                    }
                }

                // add 3 blank rows
                for (int i = 0; i <= 2; i++)
                {
                    row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                    page3TableTotalRowCount++;
                    cell = row.AddCell();
                    cell.Width = page3TableCell1Width;
                    cell = row.AddCell();
                    cell.Width = page3TableCell2Width;
                    cell = row.AddCell();
                    cell.Width = page3TableCell3_4Width;
                    cell = row.AddCell();
                    cell.Width = page3TableCell3_4Width;
                }

                // add EXPENDITURE row to table
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.AddParagraph().AppendText("EXPENDITURE\n").CharacterFormat.Bold = true;
                cell.Width = page3TableCell1Width;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;

                // add Expenditure data rows to table
                for (int i = 2; i <= page03ExpenditureWorksheet.Dimension.End.Row; i++)
                {
                    row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                    page3TableTotalRowCount++;

                    // cell 1
                    cell = row.AddCell();
                    cell.Width = page3TableCell1Width;
                    cell.AddParagraph().AppendText($"{page03ExpenditureWorksheet.Cells[i, 1].Text}");
                    page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;

                    // cell 2
                    cell = row.AddCell();
                    cell.Width = page3TableCell2Width;
                    cell.AddParagraph().AppendText($"{page03ExpenditureWorksheet.Cells[i, 2].Text}");
                    page3Table.Rows[page3TableTotalRowCount - 1].Cells[1].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;

                    // cell 3
                    cell = row.AddCell();
                    cell.Width = page3TableCell3_4Width;
                    cell.AddParagraph().AppendText($"{page03ExpenditureWorksheet.Cells[i, 3].Text}");
                    page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                    // cell 4
                    cell = row.AddCell();
                    cell.Width = page3TableCell3_4Width;
                    cell.AddParagraph().AppendText($"{page03ExpenditureWorksheet.Cells[i, 4].Text}");
                    page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                }

                // add 3 blank rows
                // for (int i = 0; i <= 2; i++)
                // {
                //     row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                //     page3TableTotalRowCount++;
                //     cell = row.AddCell();
                //     cell.Width = page3TableCell1Width;
                //     cell = row.AddCell();
                //     cell.Width = page3TableCell2Width;
                //     cell = row.AddCell();
                //     cell.Width = page3TableCell3_4Width;
                //     cell = row.AddCell();
                //     cell.Width = page3TableCell3_4Width;
                // }

                // calculate the income, minus costs:
                int revenueTotalColumnC = 0;
                for (int i = 2; i <= page03IncomeWorksheet.Dimension.End.Row; i++)
                {
                    revenueTotalColumnC = revenueTotalColumnC + page03IncomeWorksheet.Cells[i, 3].GetValue<int>();
                }
                int revenueTotalColumnD = 0;
                for (int i = 2; i <= page03IncomeWorksheet.Dimension.End.Row; i++)
                {
                    revenueTotalColumnD = revenueTotalColumnD + page03IncomeWorksheet.Cells[i, 4].GetValue<int>();
                }
                int expenseTotalColumnC = 0;
                for (int i = 2; i <= page03ExpenditureWorksheet.Dimension.End.Row; i++)
                {
                    expenseTotalColumnC = expenseTotalColumnC + page03ExpenditureWorksheet.Cells[i, 3].GetValue<int>();
                }
                int expenseTotalColumnD = 0;
                for (int i = 2; i <= page03ExpenditureWorksheet.Dimension.End.Row; i++)
                {
                    expenseTotalColumnD = expenseTotalColumnD + page03ExpenditureWorksheet.Cells[i, 4].GetValue<int>();
                }
                // Console.WriteLine($"{revenueTotalColumnC} + {expenseTotalColumnC} = {revenueTotalColumnC + expenseTotalColumnC}");
                // Console.WriteLine($"{revenueTotalColumnD} + {expenseTotalColumnD} = {revenueTotalColumnD + expenseTotalColumnD}");
                
                // profit before income tax
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Profit Before Income Tax");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{(revenueTotalColumnC + expenseTotalColumnC).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{(revenueTotalColumnD + expenseTotalColumnD).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                // calculate income tax rate
                int currentFiscalYear = page03IncomeWorksheet.Cells[1, 3].GetValue<int>();
                int lastFiscalYear = page03IncomeWorksheet.Cells[1, 4].GetValue<int>();

                double columnCTaxRate = CalculatorTool.CalculateTaxRate(thisFiscalYear: currentFiscalYear, thisFiscalYearPBTI: revenueTotalColumnC + expenseTotalColumnC);
                double columnDTaxRate = CalculatorTool.CalculateTaxRate(thisFiscalYear: lastFiscalYear, thisFiscalYearPBTI: revenueTotalColumnD + expenseTotalColumnD);
                
                // Console.WriteLine($"{currentFiscalYear} : {columnCTaxRate}");
                // Console.WriteLine($"{lastFiscalYear} : {columnDTaxRate}");

                // write Income Tax Expense
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Income Tax Expense");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate((revenueTotalColumnC + expenseTotalColumnC) * columnCTaxRate)).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate((revenueTotalColumnD + expenseTotalColumnD) * columnDTaxRate)).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                // write Profit for the Year
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Profit for the financial year");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate(revenueTotalColumnC + expenseTotalColumnC - ((revenueTotalColumnC + expenseTotalColumnC) * columnCTaxRate))).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate(revenueTotalColumnD + expenseTotalColumnD - ((revenueTotalColumnD + expenseTotalColumnD) * columnDTaxRate))).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                /*
                 * here we need to calculate the remaining fields for the rest of the form
                 * some of these calculations will follow through to the lower sections, as
                 * they will carry from ColumnD into ColumnC further down
                 */
            // int retainedEarningsLastFiscalYear,
            // int dividendPaidLastFiscalYear,
            // int dividendPaidThisFiscalYear,
                
                double profitFromColumnD = revenueTotalColumnD + expenseTotalColumnD - ((revenueTotalColumnD + expenseTotalColumnD) * columnDTaxRate);
                profitFromColumnD = profitFromColumnD + retainedEarningsLastFiscalYear;
                profitFromColumnD = profitFromColumnD - dividendPaidLastFiscalYear;
                

                // write retained earnings
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Initial retained earnings for the financial year");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate(profitFromColumnD)).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{retainedEarningsLastFiscalYear.ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                // write dividend paid
                // int dividendPaidLastFiscalYear,
                // int dividendPaidThisFiscalYear,
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Dividend Paid");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{dividendPaidThisFiscalYear.ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{dividendPaidLastFiscalYear.ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                // write retained earnings
                // int dividendPaidLastFiscalYear,
                // int dividendPaidThisFiscalYear,
                row = page3Table.AddRow(isCopyFormat: true, autoPopulateCells: false);
                page3TableTotalRowCount++;
                cell = row.AddCell();
                cell.Width = page3TableCell1Width;
                cell.AddParagraph().AppendText("Final retained earnings for the financial year");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                cell = row.AddCell();
                cell.Width = page3TableCell2Width;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                double thisFiscalYearFinalRetainedEarnings = (double)Math.Truncate(revenueTotalColumnC + expenseTotalColumnC - ((revenueTotalColumnC + expenseTotalColumnC) * columnCTaxRate)) + (double)Math.Truncate(profitFromColumnD);
                cell.AddParagraph().AppendText($"{thisFiscalYearFinalRetainedEarnings.ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[2].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                cell = row.AddCell();
                cell.Width = page3TableCell3_4Width;
                cell.AddParagraph().AppendText($"{((double)Math.Truncate(profitFromColumnD)).ToString("C", CultureInfo.CurrentCulture)}");
                page3Table.Rows[page3TableTotalRowCount - 1].Cells[3].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                IWParagraph page03EndParagraph = section03.AddParagraph();
                page03EndParagraph.AppendText($"\n\nThese notes should be read in conjunction with the attached compilation report.");
                page03EndParagraph.ApplyStyle("Page03Style02");
                page03EndParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                IWParagraph page03Footer = section03.HeadersFooters.Footer.AddParagraph();
                page03Footer.AppendText($"Page ");
                page03Footer.AppendField("Page", Syncfusion.DocIO.FieldType.FieldPage);
                page03Footer.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;


                //page3Table.Rows[page3TableTotalRowCount-1].Cells[0].Paragraphs[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                // List<string> columnNamesPage3Text = new();
                // for (int i = 2; i <= page03IncomeWorksheet.Dimension.End.Column; i++)
                // {
                //     columnNamesPage3Text.Add(page03IncomeWorksheet.Cells[1, i].Text);
                // }
                // for (int i = 0; i <= columnNamesPage3Text.Count; i++)
                // {
                //     cell = row.AddCell();
                //     cell.AddParagraph().AppendText($"{columnNamesPage3Text[i]}");
                // }

                // page3Table.ResetCells(6, page03Worksheet.Dimension.End.Column);
                // page3Table.TableFormat.IsAutoResized = true;

                // //page3Table[0, 0].AddParagraph().AppendText("INCOME\n");
                // page3Table[0, 1].AddParagraph().AppendText("NOTES\n");

                // // start the index of columnNames by skipping first 2 and switch between 1,2 or 3 columns:
                // List<string> columnNamesPage3Text = new();
                // for (int i = 1; i <= page03Worksheet.Dimension.End.Column; i++)
                // {
                //     columnNamesPage3Text.Add(page03Worksheet.Cells[1, i].Text);
                // }

                // switch (columnNamesPage3Text.Skip(2).Count())
                // {
                //     case 1:
                //         // Headers
                //         IWParagraph page3TableCol1_02 = page3Table[0, 2].AddParagraph();
                //         page3TableCol1_02.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         page3TableCol1_02.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[0]}\n$\u00A0\u00A0\u00A0\n");
                //         // SubHeaders
                //         IWParagraph page3TableCol1_10 = page3Table[1, 0].AddParagraph();
                //         page3TableCol1_10.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //         page3TableCol1_10.AppendText($"INCOME\n");
                //         // Values
                //         IWParagraph page3TableCol1_12 = page3Table[1, 2].AddParagraph();
                //         page3TableCol1_12.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         page3TableCol1_12.AppendText($"$511,376.00\n");
                //         break;
                //     case 2:
                //         // Headers
                //         IWParagraph page3TableCol2_02 = page3Table[0, 2].AddParagraph();
                //         page3TableCol2_02.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         page3TableCol2_02.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[0]}\n$\u00A0\u00A0\u00A0");
                //         IWParagraph page3TableCol2_03 = page3Table[0, 3].AddParagraph();
                //         page3TableCol2_03.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         page3TableCol2_03.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[1]}\n$\u00A0\u00A0\u00A0");
                //         // SubHeaders
                //         IWParagraph page3TableCol2_10 = page3Table[1, 0].AddParagraph();
                //         page3TableCol2_10.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //         page3TableCol2_10.AppendText("INCOME").CharacterFormat.Bold = true;
                //         page3Table[1, 0].CellFormat.TextWrap = false;
                        
                //         // Values
                //         IWParagraph page3TableCol2_20 = page3Table[2, 0].AddParagraph();
                //         page3TableCol2_20.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //         for (int row = 2; row <= page03Worksheet.Dimension.End.Row; row++)
                //         {
                //             page3TableCol2_20.AppendText($"\u00A0\u00A0{page03Worksheet.Cells[row, 1].Text}\n");
                //         }
                //         IWParagraph page3TableCol2_22 = page3Table[2, 2].AddParagraph();
                //         page3TableCol2_22.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         for (int row = 2; row <= page03Worksheet.Dimension.End.Row; row++)
                //         {
                //             page3TableCol2_22.AppendText($"{page03Worksheet.Cells[row, 3].Text}\n");
                //         }
                //         IWParagraph page3TableCol2_23 = page3Table[2, 3].AddParagraph();
                //         page3TableCol2_23.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //         for (int row = 2; row <= page03Worksheet.Dimension.End.Row; row++)
                //         {
                //             page3TableCol2_23.AppendText($"{page03Worksheet.Cells[row, 4].Text}\n");
                //         }

                //         // HEADERS 2
                //         IWParagraph page3TableCol2_40 = page3Table[4, 0].AddParagraph();
                //         page3TableCol2_40.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //         page3TableCol2_40.AppendText("EXPENDITURE").CharacterFormat.Bold = true;

                //         // Values 2
                //         IWParagraph page3TableCol2_50 = page3Table[5, 0].AddParagraph();
                        
                //         page3TableCol2_50.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //         for (int row = 2; row <= page03ExpenditureWorksheet.Dimension.End.Row; row++)
                //         {
                //             page3TableCol2_50.AppendText($"\u00A0\u00A0{page03ExpenditureWorksheet.Cells[row, 1].Text}\n");
                //         }
                //         page3Table[5, 0].CellFormat.TextWrap = false;
                        
                        
                //         package.Dispose();
                //         break;
                //     case 3:
                //         IWParagraph page3TableCol3_02 = page3Table[0, 2].AddParagraph();
                //         page3TableCol3_02.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //         page3TableCol3_02.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[0]}\n$\u00A0\u00A0\u00A0\n");
                //         IWParagraph page3TableCol3_03 = page3Table[0, 3].AddParagraph();
                //         page3TableCol3_03.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //         page3TableCol3_03.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[1]}\n$\u00A0\u00A0\u00A0\n");
                //         IWParagraph page3TableCol3_04 = page3Table[0, 4].AddParagraph();
                //         page3TableCol3_04.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //         page3TableCol3_04.AppendText($"{columnNamesPage3Text.Skip(2).ToList()[2]}\n$\u00A0\u00A0\u00A0\n");
                //         break;
                //     default:
                //         Console.WriteLine("There is X");
                //         break;
                // }

                //page3Table[1, 0].AddParagraph().AppendText("Revenue\n");
                //IWParagraph page03TableIncome = tocTable[0, 1].AddParagraph();
                //tocIncomeStatement.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //tocIncomeStatement.AppendText("3");
                #endregion Section03

                #region Section04
                // Section04 - Title Paragraph
                IWSection section04 = wordDocument.AddSection();

                // Section04 - Page Setup
                section04.PageSetup.Orientation = PageOrientation.Portrait;
                section04.PageSetup.Margins.All = 36;

                // Section04 - Paragraph Style 04 (Title)
                IWParagraphStyle secn04Style01 = wordDocument.AddParagraphStyle("Section04Style01");
                secn04Style01.ParagraphFormat.BackColor = Color.White;
                secn04Style01.ParagraphFormat.AfterSpacing = 16f;
                secn04Style01.ParagraphFormat.BeforeSpacing = 16f;
                secn04Style01.ParagraphFormat.LineSpacing = 14f;
                secn04Style01.CharacterFormat.FontName = "Times New Roman";
                secn04Style01.CharacterFormat.FontSize = 14f;
                secn04Style01.CharacterFormat.Bold = true;

                // Section04 - Title Paragraph
                IWParagraph paragraph05 = section04.AddParagraph();
                paragraph05.AppendText($"{clientName.ToUpperInvariant()}");
                paragraph05.AppendBreak(BreakType.LineBreak);
                paragraph05.AppendText(companyAbnAcn);
                paragraph05.AppendBreak(BreakType.LineBreak);
                paragraph05.AppendBreak(BreakType.LineBreak);
                paragraph05.AppendText("NOTES TO THE FINANCIAL STATEMENTS");
                paragraph05.AppendBreak(BreakType.LineBreak);
                paragraph05.AppendText("FOR THE YEAR ENDED 30 JUNE 2020");
                paragraph05.AppendBreak(BreakType.LineBreak);

                // Section04 - HR
                paragraph05.AppendText("_________________________________________________________________");
                paragraph05.ApplyStyle("Section03Style01");
                paragraph05.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                #endregion Section04

                #region Page05
                // Page05 - Title Paragraph
                IWSection page05 = wordDocument.AddSection();

                // Page05 - Page Setup
                page05.PageSetup.Orientation = PageOrientation.Portrait;
                page05.PageSetup.Margins.All = 36;

                // Page05 - Paragraph Style 01 (Title)
                IWParagraphStyle page05Style01 = wordDocument.AddParagraphStyle("Page05Style01");
                page05Style01.ParagraphFormat.BackColor = Color.White;
                page05Style01.ParagraphFormat.AfterSpacing = 16f;
                page05Style01.ParagraphFormat.BeforeSpacing = 16f;
                page05Style01.ParagraphFormat.LineSpacing = 14f;
                page05Style01.CharacterFormat.FontName = "Times New Roman";
                page05Style01.CharacterFormat.FontSize = 14f;
                page05Style01.CharacterFormat.Bold = true;

                // Page05 - Paragraph Style 02 (Body)
                IWParagraphStyle page05Style02 = wordDocument.AddParagraphStyle("Page05Style02");
                page05Style02.ParagraphFormat.BackColor = Color.White;
                page05Style02.ParagraphFormat.AfterSpacing = 14f;
                page05Style02.ParagraphFormat.BeforeSpacing = 14f;
                page05Style02.ParagraphFormat.LineSpacing = 12f;
                page05Style02.CharacterFormat.FontName = "Times New Roman";
                page05Style02.CharacterFormat.FontSize = 12f;
                page05Style02.CharacterFormat.Bold = false;

                IWParagraphStyle page05Style02Bold = wordDocument.AddParagraphStyle("Page05Style02Bold");
                page05Style02Bold.ParagraphFormat.BackColor = Color.White;
                page05Style02Bold.ParagraphFormat.AfterSpacing = 14f;
                page05Style02Bold.ParagraphFormat.BeforeSpacing = 14f;
                page05Style02Bold.ParagraphFormat.LineSpacing = 12f;
                page05Style02Bold.CharacterFormat.FontName = "Times New Roman";
                page05Style02Bold.CharacterFormat.FontSize = 12f;
                page05Style02Bold.CharacterFormat.Bold = true;

                // Page05 - Title Paragraph
                IWParagraph page05TitleParagraph = page05.AddParagraph();
                page05TitleParagraph.AppendText($"{clientName.ToUpperInvariant()}");
                page05TitleParagraph.AppendBreak(BreakType.LineBreak);
                page05TitleParagraph.AppendText(companyAbnAcn);
                page05TitleParagraph.AppendBreak(BreakType.LineBreak);
                page05TitleParagraph.AppendBreak(BreakType.LineBreak);
                page05TitleParagraph.AppendText("NOTES TO THE FINANCIAL STATEMENTS");
                page05TitleParagraph.AppendBreak(BreakType.LineBreak);
                page05TitleParagraph.AppendText("FOR THE YEAR ENDED 30 JUNE 2020");
                page05TitleParagraph.AppendBreak(BreakType.LineBreak);

                // Page05 - HR
                page05TitleParagraph.AppendText("__________________________________________________________________________");
                page05TitleParagraph.ApplyStyle("Page05Style01");
                page05TitleParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Page05 - Introduction
                IWParagraph page05IntroductionParagraph = page05.AddParagraph();
                page05IntroductionParagraph.AppendText($"The financial statements cover the business of {clientName.ToUpperInvariant()} and have been prepared to meet the needs of stakeholders " +
                    $"and to assist in the preparation of the tax return.\nComparatives are consistent with prior years, unless otherwise stated.");
                page05IntroductionParagraph.ApplyStyle("Page05Style02");

                // Page05 - Basis of Preparation (Title)
                IWParagraph page05BasisOfPreparationTitle = page05.AddParagraph();
                page05BasisOfPreparationTitle.AppendText("1.\tBasis of Preparation");
                page05BasisOfPreparationTitle.ApplyStyle("Page05Style01");

                // Page 05 - Basis of Prepartion (Text)
                IWParagraph page05BasisOfPreparationText = page05.AddParagraph();
                page05BasisOfPreparationText.ParagraphFormat.LeftIndent = 36;
                page05BasisOfPreparationText.AppendText("The Company is non reporting since there are unlikely to be any users who would rely on the general-purpose financial statements.\n\n" +
                    "The special purpose financial statements have been prepared in accordance with the significant accounting policies described below and do not comply with any Australian Accounting " +
                    "Standards unless otherwise stated.\n\n" +
                    "The financial statements have been prepared on an accrual basis and are based on historical costs modified, where applicable, by the measurement at fair value of selected " +
                    "noncurrent assets, financial assets and financial liabilities.\n\n" +
                    "Significant account policies adopted in the preparation of these financial statements are presented below and are consistent with prior reporting periods unless otherwise stated.");

                // Page05 - Summary of Significant Accounting Policies (Title)
                IWParagraph page05SummaryOfSignificantAccountingPoliciesTitle = page05.AddParagraph();
                page05SummaryOfSignificantAccountingPoliciesTitle.AppendText("2.\tSummary of Significant Accounting Policies");
                page05SummaryOfSignificantAccountingPoliciesTitle.ApplyStyle("Page05Style01");

                // Page05 - Property, Plant and Equipment (SubTitle01)
                IWParagraph page05Section02SubTitle01 = page05.AddParagraph();
                page05Section02SubTitle01.AppendText("Property, Plant and Equipment");
                page05Section02SubTitle01.ParagraphFormat.LeftIndent = 36;
                page05Section02SubTitle01.ApplyStyle("Page05Style01");

                // Page 05 - Page05 - Property, Plant and Equipment (SubText01)
                IWParagraph page05Section02SubText01 = page05.AddParagraph();
                page05Section02SubText01.ParagraphFormat.LeftIndent = 36;
                page05Section02SubText01.AppendText("Each class of property, plant and equipment is carried at cost less, where applicable, any accumulated depreciation and impairment.");

                // Page05 - Property, Plant and Equipment (SubTitle01NoteTitle01)
                IWParagraph page05Section02SubTitle01NoteTitle01 = page05.AddParagraph();
                page05Section02SubTitle01NoteTitle01.AppendText("Depreciation");
                page05Section02SubTitle01NoteTitle01.ParagraphFormat.LeftIndent = 36;
                page05Section02SubTitle01NoteTitle01.ApplyStyle("Page05Style02Bold");

                // Page05 - Property, Plant and Equipment (SubTitle01NoteText01)
                IWParagraph page05Section02SubTitle01NoteText01 = page05.AddParagraph();
                page05Section02SubTitle01NoteText01.AppendText("Property, plant and equipment excluding freehold land, is depreciated on a straight-line basis over the assets useful life to the company, " +
                    "commencing when the asset is ready for use.");
                page05Section02SubTitle01NoteText01.ParagraphFormat.LeftIndent = 36;
                page05Section02SubTitle01NoteText01.ApplyStyle("Page05Style02");
                #endregion Page05

                #region Page06
                // Page06 - Title Paragraph
                IWSection page06 = wordDocument.AddSection();

                // Page06 - Page Setup
                page06.PageSetup.Orientation = PageOrientation.Portrait;
                page06.PageSetup.Margins.All = 36;

                // Page06 - Paragraph Style 01 (Title)
                IWParagraphStyle page06Style01 = wordDocument.AddParagraphStyle("Page06Style01");
                page06Style01.ParagraphFormat.BackColor = Color.White;
                page06Style01.ParagraphFormat.AfterSpacing = 16f;
                page06Style01.ParagraphFormat.BeforeSpacing = 16f;
                page06Style01.ParagraphFormat.LineSpacing = 14f;
                page06Style01.CharacterFormat.FontName = "Times New Roman";
                page06Style01.CharacterFormat.FontSize = 14f;
                page06Style01.CharacterFormat.Bold = true;

                // Page06 - Paragraph Style 02 (Body)
                IWParagraphStyle page06Style02 = wordDocument.AddParagraphStyle("Page06Style02");
                page06Style02.ParagraphFormat.BackColor = Color.White;
                page06Style02.ParagraphFormat.AfterSpacing = 14f;
                page06Style02.ParagraphFormat.BeforeSpacing = 14f;
                page06Style02.ParagraphFormat.LineSpacing = 12f;
                page06Style02.CharacterFormat.FontName = "Times New Roman";
                page06Style02.CharacterFormat.FontSize = 12f;
                page06Style02.CharacterFormat.Bold = false;

                IWParagraphStyle page06Style02Bold = wordDocument.AddParagraphStyle("Page06Style02Bold");
                page06Style02Bold.ParagraphFormat.BackColor = Color.White;
                page06Style02Bold.ParagraphFormat.AfterSpacing = 14f;
                page06Style02Bold.ParagraphFormat.BeforeSpacing = 14f;
                page06Style02Bold.ParagraphFormat.LineSpacing = 12f;
                page06Style02Bold.CharacterFormat.FontName = "Times New Roman";
                page06Style02Bold.CharacterFormat.FontSize = 12f;
                page06Style02Bold.CharacterFormat.Bold = true;

                // Page06 - Title Paragraph
                IWParagraph page06TitleParagraph = page06.AddParagraph();
                page06TitleParagraph.AppendText($"{clientName.ToUpperInvariant()}");
                page06TitleParagraph.AppendBreak(BreakType.LineBreak);
                page06TitleParagraph.AppendText(companyAbnAcn);
                page06TitleParagraph.AppendBreak(BreakType.LineBreak);
                page06TitleParagraph.AppendBreak(BreakType.LineBreak);
                page06TitleParagraph.AppendText("DIRECTOR'S DECLARATION");
                page06TitleParagraph.AppendBreak(BreakType.LineBreak);

                // Page06 - HR
                page06TitleParagraph.AppendText("__________________________________________________________________________");
                page06TitleParagraph.ApplyStyle("Page06Style01");
                page06TitleParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                #endregion Page06

                #region Page07
                // Page07 - Title Paragraph
                IWSection page07 = wordDocument.AddSection();

                // Page07 - Page Setup
                page07.PageSetup.Orientation = PageOrientation.Portrait;
                page07.PageSetup.Margins.All = 36;

                // Page07 - Paragraph Style 01 (Title)
                IWParagraphStyle page07Style01 = wordDocument.AddParagraphStyle("Page07Style01");
                page07Style01.ParagraphFormat.BackColor = Color.White;
                page07Style01.ParagraphFormat.AfterSpacing = 16f;
                page07Style01.ParagraphFormat.BeforeSpacing = 16f;
                page07Style01.ParagraphFormat.LineSpacing = 14f;
                page07Style01.CharacterFormat.FontName = "Times New Roman";
                page07Style01.CharacterFormat.FontSize = 14f;
                page07Style01.CharacterFormat.Bold = true;

                // Page07 - Paragraph Style 02 (Body)
                IWParagraphStyle page07Style02 = wordDocument.AddParagraphStyle("Page07Style02");
                page07Style02.ParagraphFormat.BackColor = Color.White;
                page07Style02.ParagraphFormat.AfterSpacing = 0f;
                page07Style02.ParagraphFormat.BeforeSpacing = 0f;
                page07Style02.ParagraphFormat.LineSpacing = 12f;
                page07Style02.CharacterFormat.FontName = "Times New Roman";
                page07Style02.CharacterFormat.FontSize = 12f;
                page07Style02.CharacterFormat.Bold = false;

                IWParagraphStyle page07Style02Bold = wordDocument.AddParagraphStyle("Page07Style02Bold");
                page07Style02Bold.ParagraphFormat.BackColor = Color.White;
                page07Style02Bold.ParagraphFormat.AfterSpacing = 6f;
                page07Style02Bold.ParagraphFormat.BeforeSpacing = 0f;
                page07Style02Bold.ParagraphFormat.LineSpacing = 12f;
                page07Style02Bold.CharacterFormat.FontName = "Times New Roman";
                page07Style02Bold.CharacterFormat.FontSize = 12f;
                page07Style02Bold.CharacterFormat.Bold = true;

                // Page07 - Title Paragraph
                IWParagraph page07TitleParagraph = page07.AddParagraph();
                page07TitleParagraph.AppendText($"{clientName.ToUpperInvariant()}");
                page07TitleParagraph.AppendBreak(BreakType.LineBreak);
                page07TitleParagraph.AppendText(companyAbnAcn);
                page07TitleParagraph.AppendBreak(BreakType.LineBreak);
                page07TitleParagraph.AppendBreak(BreakType.LineBreak);
                page07TitleParagraph.AppendText("DIRECTOR'S DECLARATION");
                page07TitleParagraph.AppendBreak(BreakType.LineBreak);

                // Page07 - HR
                page07TitleParagraph.AppendText("__________________________________________________________________________");
                page07TitleParagraph.ApplyStyle("Page07Style01");
                page07TitleParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Page07 - Entry text
                IWParagraph page07EntryText = page07.AddParagraph();
                page07EntryText.AppendText("The director of the company declares that:\n");
                page07EntryText.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - Declaration01
                IWParagraph page07Declaration01 = page07.AddParagraph();
                page07Declaration01.AppendText("1.\tThe financial statements and notes, as set out of pages 1 to 6, for the year ended 30 June 2020 are in accordance with the Corporations Act 2001 and:\n");
                page07Declaration01.ParagraphFormat.FirstLineIndent = -36;
                page07Declaration01.ParagraphFormat.LeftIndent = 36;
                page07Declaration01.ParagraphFormat.Keep = true;
                page07Declaration01.ParagraphFormat.KeepFollow = true;
                page07Declaration01.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - Declaration01PointA
                IWParagraph page07Declaration01PointA = page07.AddParagraph();
                page07Declaration01PointA.AppendText("(a)\tcomply with Accounting Standards, which, as stated in accounting policy Note 1 to the financial statements, constitutes explicit and unreserved " +
                    "compliance with International Financial Reporting Standards (IFRS); and\n");
                page07Declaration01PointA.ParagraphFormat.FirstLineIndent = -20;
                page07Declaration01PointA.ParagraphFormat.LeftIndent = 36;
                page07Declaration01PointA.ParagraphFormat.Keep = true;
                page07Declaration01PointA.ParagraphFormat.KeepFollow = true;
                page07Declaration01PointA.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - Declaration01PointB
                IWParagraph page07Declaration01PointB = page07.AddParagraph();
                page07Declaration01PointB.AppendText("(b)\tgive a true and fair view of the financial position and performance of the company.\n");
                page07Declaration01PointB.ParagraphFormat.FirstLineIndent = -20;
                page07Declaration01PointB.ParagraphFormat.LeftIndent = 36;
                page07Declaration01PointB.ParagraphFormat.Keep = true;
                page07Declaration01PointB.ParagraphFormat.KeepFollow = true;
                page07Declaration01PointB.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - Declaration02
                IWParagraph page07Declaration02 = page07.AddParagraph();
                page07Declaration02.AppendText("2.\tIn the director's opnion, there are reasonable grounds to believe that the company will be able to pay its debts as and when they become due and payable.\n");
                page07Declaration02.ParagraphFormat.FirstLineIndent = -36;
                page07Declaration02.ParagraphFormat.LeftIndent = 36;
                page07Declaration02.ParagraphFormat.Keep = true;
                page07Declaration02.ParagraphFormat.KeepFollow = true;
                page07Declaration02.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - DeclarationText
                IWParagraph page07DeclarationText = page07.AddParagraph();
                page07DeclarationText.AppendText("This declaration is made in accordance with a resolution of the director.\n");
                page07DeclarationText.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - DeclarationSignatureLine
                IWParagraph page07DeclarationSignatureLine = page07.AddParagraph();
                page07DeclarationSignatureLine.AppendText("Director:\t___________________________________________");
                page07DeclarationSignatureLine.ApplyStyle("Page07Style02Bold");

                // Page07 - DeclarationSignatureName
                IWParagraph page07DeclarationSignatureName = page07.AddParagraph();
                page07DeclarationSignatureName.AppendText("\t\tHabib Zaiter\n\n");
                page07DeclarationSignatureName.ApplyStyle(BuiltinStyle.Normal);

                // Page07 - DeclarationDated
                IWParagraph page07DeclarationDated = page07.AddParagraph();
                IWTextRange datedThisTextRange = new WTextRange(wordDocument);
                datedThisTextRange.Text = "Dated this ";
                datedThisTextRange.CharacterFormat.FontSize = 12f;
                datedThisTextRange.CharacterFormat.FontName = "Times New Roman";
                datedThisTextRange.CharacterFormat.Bold = true;

                IWTextRange calendarDayTextRange = new WTextRange(wordDocument);
                calendarDayTextRange.Text = "30th";
                calendarDayTextRange.CharacterFormat.FontSize = 12f;
                calendarDayTextRange.CharacterFormat.FontName = "Times New Roman";
                calendarDayTextRange.CharacterFormat.Bold = false;

                IWTextRange dayOfTextRange = new WTextRange(wordDocument);
                dayOfTextRange.Text = " day of ";
                dayOfTextRange.CharacterFormat.FontSize = 12f;
                dayOfTextRange.CharacterFormat.FontName = "Times New Roman";
                dayOfTextRange.CharacterFormat.Bold = true;

                IWTextRange calendarMonthTextRange = new WTextRange(wordDocument);
                calendarMonthTextRange.Text = "November 2021";
                calendarMonthTextRange.CharacterFormat.FontSize = 12f;
                calendarMonthTextRange.CharacterFormat.FontName = "Times New Roman";
                calendarMonthTextRange.CharacterFormat.Bold = false;

                page07DeclarationDated.Items.Add(datedThisTextRange);
                page07DeclarationDated.Items.Add(calendarDayTextRange);
                page07DeclarationDated.Items.Add(dayOfTextRange);
                page07DeclarationDated.Items.Add(calendarMonthTextRange);
                #endregion Page07

                #region Page08
                // Page08 - Compilation Report
                IWSection page08 = wordDocument.AddSection();

                // Page08 - Page Setup
                page08.PageSetup.Orientation = PageOrientation.Portrait;
                page08.PageSetup.Margins.All = 36;

                // Page08 - Paragraph Style 01 (Title)
                IWParagraphStyle page08Style01 = wordDocument.AddParagraphStyle("Page08Style01");
                page08Style01.ParagraphFormat.BackColor = Color.White;
                page08Style01.ParagraphFormat.AfterSpacing = 16f;
                page08Style01.ParagraphFormat.BeforeSpacing = 16f;
                page08Style01.ParagraphFormat.LineSpacing = 14f;
                page08Style01.CharacterFormat.FontName = "Times New Roman";
                page08Style01.CharacterFormat.FontSize = 14f;
                page08Style01.CharacterFormat.Bold = true;

                // Page08 - Paragraph Style 02 (Body)
                IWParagraphStyle page08Style02 = wordDocument.AddParagraphStyle("Page08Style02");
                page08Style02.ParagraphFormat.BackColor = Color.White;
                page08Style02.ParagraphFormat.AfterSpacing = 0f;
                page08Style02.ParagraphFormat.BeforeSpacing = 0f;
                page08Style02.ParagraphFormat.LineSpacing = 12f;
                page08Style02.CharacterFormat.FontName = "Times New Roman";
                page08Style02.CharacterFormat.FontSize = 12f;
                page08Style02.CharacterFormat.Bold = false;

                // Page08 - Paragraph Style 02 (Body-Bold)
                IWParagraphStyle page08Style02Bold = wordDocument.AddParagraphStyle("Page08Style02Bold");
                page08Style02Bold.ParagraphFormat.BackColor = Color.White;
                page08Style02Bold.ParagraphFormat.AfterSpacing = 6f;
                page08Style02Bold.ParagraphFormat.BeforeSpacing = 0f;
                page08Style02Bold.ParagraphFormat.LineSpacing = 12f;
                page08Style02Bold.CharacterFormat.FontName = "Times New Roman";
                page08Style02Bold.CharacterFormat.FontSize = 12f;
                page08Style02Bold.CharacterFormat.Bold = true;

                // Page08 - Title Paragraph
                IWParagraph page08TitleParagraph = page08.AddParagraph();

                page08TitleParagraph.AppendText($"{clientName.ToUpperInvariant()}");
                page08TitleParagraph.AppendBreak(BreakType.LineBreak);
                page08TitleParagraph.AppendText(companyAbnAcn);
                page08TitleParagraph.AppendBreak(BreakType.LineBreak);
                page08TitleParagraph.AppendBreak(BreakType.LineBreak);
                page08TitleParagraph.AppendText("COMPILATION REPORT");
                page08TitleParagraph.AppendBreak(BreakType.LineBreak);

                // Page08 - HR
                page08TitleParagraph.AppendText("__________________________________________________________________________");
                page08TitleParagraph.ApplyStyle("Page08Style01");
                page08TitleParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                // Page08 - Entry text
                IWParagraph page08EntryText = page08.AddParagraph();
                page08EntryText.AppendText($"We have compiled the accompanying special purpose financial statements of {clientName.ToUpperInvariant()} which comprise the balance sheet as at 30 June 2020, " +
                    $"and the income statement for the year then ended, a summary of significant accounting policies and other explanatory notes.\n\n" +
                    $"The specific purpose for which the special purpose financial statements have been prepared is set out in the notes to the accounts.\n");
                page08EntryText.ApplyStyle(BuiltinStyle.Normal);

                // Page08 - Director's Responsibility (Title)
                IWParagraph page08DirectorsResponsibilityTitle = page08.AddParagraph();
                page08DirectorsResponsibilityTitle.AppendText($"The Responsibility of the Director");
                page08DirectorsResponsibilityTitle.ApplyStyle("Page08Style02Bold");

                // Page08 - Director's Responsibility (Text)
                IWParagraph page08DirectorsResponsibilityText = page08.AddParagraph();
                page08DirectorsResponsibilityText.AppendText($"The director of {clientName.ToUpperInvariant()} is solely responsible for the information contained in the special purpose financial " +
                    $"statements, the reliability, accuracy and completeness of the information and for the determination that the basis of accounting used is appropriate to meet the needs and for " +
                    $"the purpose that the financial statements were prepared.\n");
                page08DirectorsResponsibilityText.ApplyStyle("Page08Style02");

                // Page08 - Our Responsibility (Title)
                IWParagraph page08OurResponsibilityTitle = page08.AddParagraph();
                page08OurResponsibilityTitle.AppendText($"Our Responsibility");
                page08OurResponsibilityTitle.ApplyStyle("Page08Style02Bold");

                // Page08 - Our Responsibility (Text)
                IWParagraph page08OurResponsibilityText = page08.AddParagraph();
                page08OurResponsibilityText.AppendText($"On the basis of the information provided by the director, we have compiled the accompanying special purpose financial statements in " +
                    $"accordance with the basis of accounting as described in the notes to the financial statements and APES 315: Compilation of Financial Information.\n\n" +
                    $"We have applied professional expertise in accounting and financial reporting to compile these financial statements in accordance with the basis of accounting described " +
                    $"in the notes to the financial statements. We have complied with the relevant ethical requirements of APES 110: Code of Ethics for Professional Accountants.\n");
                page08OurResponsibilityText.ApplyStyle("Page08Style02");

                // Page08 - Assurance Disclaimer (Title)
                IWParagraph page08AssuranceDisclaimerTitle = page08.AddParagraph();
                page08AssuranceDisclaimerTitle.AppendText($"Assurance Disclaimer");
                page08AssuranceDisclaimerTitle.ApplyStyle("Page08Style02Bold");

                // Page08 - Assurance Disclaimer (Text)
                IWParagraph page08AssuranceDisclaimerText = page08.AddParagraph();
                page08AssuranceDisclaimerText.AppendText($"Since a compilation engagement is not an assurance of engagement, we are not required to verify the reliability, accuracy or completeness " +
                    $"of the information provided to us by management to compile these financial statements. Accordingly, we do not express an audit opinion or a review conclusion on these " +
                    $"financial statements.\n\n" +
                    $"The special purpose financial statements were compiled for the benefit of the director who is responsible for the reliability, accuracy and completeness of the information used to " +
                    $"compile them. We do not accept responsibility to any other person for the contents of the special purpose financial statements.\n\n");
                page08AssuranceDisclaimerText.ApplyStyle("Page08Style02");

                // Page08 - Name of Firm
                IWParagraph page08NameOfFirm = page08.AddParagraph();
                IWTextRange nameOfFirmTextRange = new WTextRange(wordDocument);
                nameOfFirmTextRange.Text = "Name of Firm:\t";
                nameOfFirmTextRange.CharacterFormat.FontSize = 12f;
                nameOfFirmTextRange.CharacterFormat.FontName = "Times New Roman";
                nameOfFirmTextRange.CharacterFormat.Bold = true;

                IWTextRange nameFirmNameTextRange = new WTextRange(wordDocument);
                nameFirmNameTextRange.Text = "Business Accounting and Tax Solutions\n\n";
                nameFirmNameTextRange.CharacterFormat.FontSize = 12f;
                nameFirmNameTextRange.CharacterFormat.FontName = "Times New Roman";
                nameFirmNameTextRange.CharacterFormat.Bold = false;

                page08NameOfFirm.Items.Add(nameOfFirmTextRange);
                page08NameOfFirm.Items.Add(nameFirmNameTextRange);

                
                // Page08 - Name of Partner
                IWParagraph page08NameOfPartner = page08.AddParagraph();
                IWTextRange nameNameOfPartnerTextRange = new WTextRange(wordDocument);
                nameNameOfPartnerTextRange.Text = "Name of Partner:\t___________________________________________\n";
                nameNameOfPartnerTextRange.CharacterFormat.FontSize = 12f;
                nameNameOfPartnerTextRange.CharacterFormat.FontName = "Times New Roman";
                nameNameOfPartnerTextRange.CharacterFormat.Bold = true;

                IWTextRange nameNameOfPartnerNameTextRange = new WTextRange(wordDocument);
                nameNameOfPartnerNameTextRange.Text = "\t\t\tFaranak Farnosh\n";
                nameNameOfPartnerNameTextRange.CharacterFormat.FontSize = 12f;
                nameNameOfPartnerNameTextRange.CharacterFormat.FontName = "Times New Roman";
                nameNameOfPartnerNameTextRange.CharacterFormat.Bold = false;

                page08NameOfPartner.Items.Add(nameNameOfPartnerTextRange);
                page08NameOfPartner.Items.Add(nameNameOfPartnerNameTextRange);

                // Page08 - Address of Firm
                IWParagraph page08AddressOfFirm = page08.AddParagraph();
                IWTextRange AddressOfFirmTextRange = new WTextRange(wordDocument);
                AddressOfFirmTextRange.Text = "Address:\t\t";
                AddressOfFirmTextRange.CharacterFormat.FontSize = 12f;
                AddressOfFirmTextRange.CharacterFormat.FontName = "Times New Roman";
                AddressOfFirmTextRange.CharacterFormat.Bold = true;

                IWTextRange firmAddressTextRange = new WTextRange(wordDocument);
                firmAddressTextRange.Text = "52 Benaroon Ave, ST IVES, NSW, 2755\n";
                firmAddressTextRange.CharacterFormat.FontSize = 12f;
                firmAddressTextRange.CharacterFormat.FontName = "Times New Roman";
                firmAddressTextRange.CharacterFormat.Bold = false;

                page08AddressOfFirm.Items.Add(AddressOfFirmTextRange);
                page08AddressOfFirm.Items.Add(firmAddressTextRange);

                // Page08 - CompilationDated
                IWParagraph page08CompilationDated = page08.AddParagraph();
                page08CompilationDated.Items.Add(datedThisTextRange);
                page08CompilationDated.Items.Add(calendarDayTextRange);
                page08CompilationDated.Items.Add(dayOfTextRange);
                page08CompilationDated.Items.Add(calendarMonthTextRange);
                #endregion Page08

                // Saves the Word document to MemoryStream
                MemoryStream stream = new MemoryStream();
                wordDocument.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                stream.Position = 0;
                return stream;
            }
        }
    }
}
