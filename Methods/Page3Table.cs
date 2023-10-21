using Syncfusion.DocIO.DLS;

namespace FinReportBuilderCLI.Methods
{
	public class Page3Table
	{
		public static IWTable CreateTable(WordDocument document)
		{
			IWTable table = new WTable(document);
			table.ResetCells(3, 2);

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 2; col++)
                {
                    IWTextRange textRange = table[row, col].AddParagraph().AppendText("Cell Text");
                    textRange.CharacterFormat.FontName = "Arial";
                    textRange.CharacterFormat.FontSize = 10;
                }
            }

            return table;
        }
	}
}

