using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace documorph;

public sealed class TableModel(Table table, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
{
    private readonly Table _table = table;
    private readonly IEnumerable<HyperlinkRelationship> _hyperlinkRelationships = hyperlinkRelationships;

    public void AppendMarkdown(StringBuilder builder)
    {
        var isFirstRow = true;

        foreach (var row in _table.Elements<TableRow>())
        {
            if (!isFirstRow)
                builder.AppendLine();

            foreach (var cell in row.Elements<TableCell>())
            {
                builder.Append('|');
                foreach (var element in cell.Elements())
                {
                    switch (element)
                    {
                        case Paragraph paragraph:
                            new ParagraphModel(paragraph, null, _hyperlinkRelationships, []).AppendMarkdown(builder, false, true);
                            break;
                    }
                }
            }
            builder.Append('|');

            if (isFirstRow)
            {
                builder.AppendLine();
                builder.Append('|');
                foreach (var cell in row.Elements<TableCell>())
                    builder.Append("---|");
            }

            isFirstRow = false;
        }
        
        builder.AppendLine();
    }
}