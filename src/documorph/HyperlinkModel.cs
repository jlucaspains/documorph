
namespace documorph;

public sealed class HyperlinkModel(IEnumerable<RunModel> runs, string? uri): IParagraphChildModel
{
    public IEnumerable<RunModel> Runs { get; private set; } = runs;
    public string? Uri { get; set; } = uri;

    public static HyperlinkModel FromHyperlink(Hyperlink hyperlink, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
    {
        var runs = hyperlink.ChildElements.OfType<Run>()
            .Select(run => RunModel.FromRun(run, []))
            .ToList();

        return new HyperlinkModel(runs, GetUri(hyperlink, hyperlinkRelationships));
    }

    private static string GetUri(Hyperlink hyperlink, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
    {
        var hyperlinkId = hyperlink.Id?.Value;
        return hyperlinkRelationships
            .FirstOrDefault(item => item.Id == hyperlinkId)?.Uri.ToString()
            ?? string.Empty;
    }
}
