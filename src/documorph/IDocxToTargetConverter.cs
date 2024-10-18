using documorph;

/// <summary>
/// Defines a contract for converting DOCX files to a target format.
/// </summary>
public interface IDocxToTargetConverter
{
    /// <summary>
    /// /// Processes the DOCX file and converts it to the target format.
    /// </summary>
    /// <returns>
    /// A tuple containing the result as a string and an enumerable collection of <see cref="MediaModel"/> objects.
    /// </returns>
    (string Result, IEnumerable<MediaModel> Media) Process();
}