namespace documorph;

public sealed class HyperlinkModel(Hyperlink hyperlink, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
{
    public void AppendMarkdown(StringBuilder builder)
    {
        var hyperlinkId = hyperlink.Id?.Value;
        var hyperlinkUrl = hyperlinkRelationships
            .FirstOrDefault(item => item.Id == hyperlinkId)?.Uri.ToString()
            ?? string.Empty;

        var internalBuiler = new StringBuilder();

        foreach (var element in hyperlink.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    new RunModel(run, []).AppendMarkdown(internalBuiler);
                    break;
            }
        }

        var markdown = internalBuiler.Replace("[", "\\[")
            .Replace("]", "\\]")
            .Replace("(", "\\(")
            .Replace(")", "\\)")
            .ToString();

        if (!string.IsNullOrEmpty(hyperlinkUrl))
            markdown = $"[{markdown}]({hyperlinkUrl})";

        builder.Append(markdown);
    }
}
