using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace FinReportBuilderCLI.Services
{
    public class FinancialReportService
    {
        public MemoryStream CreateFinancialReportForYearEnded(
            string clientName,
            string? abn,
            string? acn)
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

                // Section03 - Title Paragraph
                IWParagraph paragraph04 = section03.AddParagraph();
                paragraph04.AppendBreak(BreakType.LineBreak);
                paragraph04.AppendBreak(BreakType.LineBreak);
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
                paragraph04.AppendText("_________________________________________________________________");
                paragraph04.ApplyStyle("Section03Style01");
                paragraph04.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                #endregion Section03

                // Saves the Word document to MemoryStream
                MemoryStream stream = new MemoryStream();
                wordDocument.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                stream.Position = 0;
                return stream;
            }
        }
    }
}
