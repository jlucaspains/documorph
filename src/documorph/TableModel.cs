namespace documorph;

public sealed class TableModel(IEnumerable<IEnumerable<ParagraphModel>> rows): IDocumentChildren
{
    public IEnumerable<IEnumerable<ParagraphModel>> Rows { get; private set; } = rows;

    public static TableModel FromTable(Table table, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
    {
        var rows = new List<IEnumerable<ParagraphModel>>();
        foreach (var row in table.Elements<TableRow>())
        {
            var rowCells = new List<ParagraphModel>();
            foreach (var cell in row.Elements<TableCell>())
            {
                foreach (var element in cell.Elements())
                {
                    switch (element)
                    {
                        case Paragraph paragraph:
                            rowCells.Add(ParagraphModel.FromParagraph(paragraph, null, hyperlinkRelationships, []));
                            break;
                    }
                }
            }
            rows.Add(rowCells);
        }

        return new TableModel(rows);
    }
}