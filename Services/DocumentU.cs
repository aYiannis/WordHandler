using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordHandler.Entities;
using static Model.Entities.NamedEntities.Condition;

namespace WordHandler.Services;

internal static class DocumentU
{
    public static Paragraph BuildParagraph(string value) => new(BuildRun(value));
    public static Run BuildRun(string value) => new(new Text(value));
    public static IEnumerable<Paragraph> BuildTitledListItem(TitledContent p, ParagraphProperties? props) {
        props = props is null ? new() : ((ParagraphProperties)props.CloneNode(true));
        yield return new Paragraph(props, new Run(new RunProperties(new Bold()), new Text(p.Title)));
        foreach (var child in p.Content) {
            props = props is null ? new() : ((ParagraphProperties)props.CloneNode(true));
            yield return new Paragraph(props, new Run(new Text(child)));
        }
	}

    public static Run? WriteAfterBookmark(BookmarkStart bookmarkStart, string value) {
        var runBefore = bookmarkStart.PreviousSibling<Run>() ?? throw new Exception("No run found!");

        // Clone the RunProperties of the run before the bookmark
        var runPropsClone = (RunProperties)runBefore.RunProperties?.CloneNode(true)!;

        // Create a new run with the cloned RunProperties
        var newRun = new Run(runPropsClone, new Text(value));


        // Find the child paragraph of the bookmark
        var parentParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
        // Insert the new run after the run before the bookmark
        return parentParagraph?.InsertAfter(newRun, runBefore);
    }

    public static void AppendGuildelines(Body body, BookmarkStart bookmarkStart, IEnumerable<TitledContent> guidelines)
    {
		OpenXmlElement rc = bookmarkStart.RootChild("No bookmark start rc level found!");
        if (rc is not Paragraph p) throw new InvalidCastException("The root child of the guildelines was not of type Paragraph.");
        ParagraphProperties? props = p.ParagraphProperties;
        OpenXmlElement current = rc;
        foreach (var guideline in guidelines) {
            foreach (var next in BuildTitledListItem(guideline, props)) {
				body.InsertAfter(next, current);
				current = next;
			}
		}
    }

    public static OpenXmlElement RootChild(this OpenXmlElement? element, string message) => element.FindAncestorWithParentOfType<Body>()
        ?? throw new Exception(message);
    public static OpenXmlElement? FindAncestorWithParentOfType<T>(this OpenXmlElement? element) where T : OpenXmlElement
    {
        if (element is null) return null;

        var parent = element.Parent;
        do {
            if (parent is T)
                return element;
            if (parent is null)
                return null;

            element = parent;
            parent = parent.Parent;
        } while (true);
    }

    public static void AddDailyWorkoutRoutine(Table table, Workout[] session)
    {
        var fr = table.Descendants<TableRow>().First();
        var cells = fr.Descendants<Text>().ToArray();
        cells[0].Text = "1";
        cells[1].Text = session[0].Exercise.ToString();
        cells[2].Text = session[0].RepsToString();

        // Assume rowData.Columns is a collection of column data for the row
        for (var iw = 1; iw < session.Length; iw++)
        {
            // Create a new ft row
            var tr = (TableRow)fr.CloneNode(true);
            var vals = tr.Descendants<Text>().ToArray();

            var workout = session[iw];
            vals[0].Text = (iw + 1).ToString();
            vals[1].Text = workout.Exercise.Name;
            vals[2].Text = workout.RepsToString();

            // Append the ft row to the ft
            table.Append(tr);
        }
    }
}
