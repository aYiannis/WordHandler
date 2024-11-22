using static WordHandler.Services.DocumentU;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Model.Entities;

using Model.Ancillary;
using DocumentFormat.OpenXml;
using Common;
using System.Text.Json;
using Ancillary.Extensions;
using System.Runtime.Serialization;

namespace WordHandler.Services;

public class DocumentCreator {
	public static void Create(ExportOrder order, string filePath) {
		// Copy template to destination path
		File.Copy(Paths.Templates, filePath, true);

		using var wordDocument = WordprocessingDocument.Open(filePath, true);
		var mainPart = wordDocument.MainDocumentPart ?? throw new NullReferenceException("Main document part was null!");
		var rootElement = mainPart.RootElement ?? throw new NullReferenceException("Root element!");
		var body = mainPart.Document.Body ?? throw new NullReferenceException("The body element of the template was null!");

		var tables = mainPart.Document.Body?.Elements<Table>().ToList() ?? throw new Exception("No resistanceTable into the resistanceTable.");
		// Assuming that the 1st rootChildTable in the template document is the resistance rootChildTable
		var resistanceTable = tables[0];
		// Assuming that the 2nd rootChildTable in the template document is the cardio rootChildTable
		var cardioTable = tables[1];

		#region Write Bookmarks
		var bookmarksStarts = FindAllElement<BookmarkStart>().ToDictionary(bm => bm.Name?.Value!);
		var bookmarksEnds = FindAllElement<BookmarkEnd>().ToDictionary(bmEnd => bmEnd?.Id?.Value!);

		if (order.Expires is not null)
			writeAfterBookmark("expiration", order.Expires.Value.ToShortDateString());
		
		if (order.User is null) {
			if (bookmarksStarts.TryGetValue("user", out var bm)) {
				var bmStartRoot = bm.RootChild("User bookmark-start not found!");
				var bmId = bm?.Id?.Value ?? throw new Exception("Bookmark-start has no id attribute!");
				OpenXmlElement? bmEndRoot = null;
				if (bookmarksEnds.TryGetValue(bmId, out var bmEnd))
					bmEndRoot = bmEnd.RootChild("User bookmark-end not found!");
				bmStartRoot.Remove();
				if (bmEndRoot != bmStartRoot)
					bmEndRoot?.Remove();
		}
		} else {
			writeAfterBookmark("user", order.User.Name);
		}

		writeAfterBookmark("trainee", order.Trainee.FullName);
		writeAfterBookmark("issuedAt", $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff (ddd)}");
		writeAfterBookmark("documentId", order.Id.ToString());
		#endregion

		#region Write Resistance
		if (tables.Count < 2) throw new Exception("Too few tables in the document.");

		var fth = tables[0].PreviousSibling<Paragraph>();
		if (order.DailyWorkoutRoutine is null || order.DailyWorkoutRoutine.Length == 0)
			// TODO: remove the header and the sets all the block (EDGE CASE)
			resistanceTable.Remove();
		else {
			if (order.Resistance is MultisetProtocol mult) {
				if (mult.Sets > 1) {
					var bookmarkStart = bookmarksStarts["setPlural"];
					WriteAfterBookmark(bookmarkStart, "s");
				}
				writeAfterBookmark("sets", mult.Sets.ToString());
			} else {
				bookmarksStarts["sets"].Remove();
			}

			var daysCount = order.DailyWorkoutRoutine.Length;
			// Create required tables for each day
			Table ot = resistanceTable;
			for (var iday = 1; iday < daysCount; iday++) {
				var t = (Table)resistanceTable.CloneNode(true);
				tables.Insert(iday, t);
				body.InsertAfter(t, ot);
				if (fth is not null) {
					var th = (Paragraph)fth.CloneNode(true);
					var textNode = th.Descendants<Text>().FirstOrDefault();
					if (textNode is not null) textNode.Text = "Μέρος " + (iday + 1);
					body.InsertBefore(th, t);
				}
				ot = t;
			}

			for (var iday = 0; iday < daysCount; iday++)
				AddDailyWorkoutRoutine(tables[iday], order.DailyWorkoutRoutine[iday]);
		}
		#endregion

		#region Write Cardio
		var trs = cardioTable.Descendants<TableRow>().ToList();
		var cardio = order.Cardio;
		if (cardio is null) {
			removeSectionByBookmarkName("cardioSection");
			removeSectionByBookmarkName("borgDescriptionSection");
		} else {
			writeOrRemove(trs[0], cardio.Wormup);
			writeOrRemove(trs[1], cardio.Main, cardio.RPE);
			writeOrRemove(trs[2], cardio.Cooldown);
			void writeOrRemove(TableRow tr, double value, RPE? rpe = null) {
				if (value > 0) {
					var ps = tr.Descendants<Paragraph>().ToArray();
					var sval = value == 1 ? "1 λεπτό"
						: value.ToString(System.Globalization.CultureInfo.CurrentCulture) + " λεπτά";
					ps[1].Append(BuildRun(sval));
					if (rpe is not null) ps[2].Append(BuildRun($"{rpe.Min}-{rpe.Max} Borg RPE"));
				} else {
					tr.Remove();
				}
			}
			if (cardio.RPE is null)
				removeSectionByBookmarkName("borgDescriptionSection");
		}
		#endregion

		// Set core properties
		var coreProperties = wordDocument.PackageProperties;
		if (coreProperties is not null) {
			coreProperties.Title = "Gym Schedule";
			coreProperties.Created = DateTime.Now;
			coreProperties.Modified = DateTime.Now;
			coreProperties.Subject = "A schedule for gym.";
			coreProperties.Description += $"\n{ORDER_ID_IDENTIFIER}{order.Id}" +
				$"{MARKER_START}{order.ToLine()}{MARKER_END}";
			coreProperties.Keywords = "resistance-do,gym-schedule";
		}

		// Add guidlines
		if (bookmarksStarts.TryGetValue("guidelines", out var glBm)) {
			var gls = order.Trainee.Conditions.Select(c => c.Guidelines!).Where(gl => gl is not null);
			if (gls.Any()) AppendGuildelines(body, glBm, gls);
		}

		// Remove sections that are not required
		if (!order.Trainee.Conditions.Any(c => c.RequiresPainScale == true))
			removeSectionByBookmarkName("borgPainSection");
		
		wordDocument.Save();

		// Methods
		void writeAfterBookmark(string bmKey, string value) {
			if (bookmarksStarts.TryGetValue(bmKey, out var bm))
				WriteAfterBookmark(bm, value);
		}

		IEnumerable<T> FindAllElement<T>() where T:OpenXmlElement {
			// Get the elements from the body
			var @enum = rootElement.Descendants<T>();
			// Get @enum from header parts
			foreach (var headerPart in mainPart.HeaderParts) {
				if (headerPart.RootElement is not null)
					@enum = @enum.Union(headerPart.RootElement.Descendants<T>());
			}
			// Get @enum from footer parts
			foreach (var footerPart in mainPart.FooterParts) {
				if (footerPart.RootElement is not null)
					@enum = @enum.Union(footerPart.RootElement.Descendants<T>());
			}
			return @enum;
		}

		/// <summary> Removes the header and the following table based on the bookmark. </summary>
		void removeSectionByBookmarkName(string bmName) {
			if (!bookmarksStarts.TryGetValue(bmName, out var bmStart)) return;
			OpenXmlElement rc = bmStart.RootChild("Failed to find the 2nd generation ansestor of the bookmark start with name " + bmName);
			rc.NextSibling<Table>()?.Remove();
			if (rc is Paragraph header)
				header.Remove();
		}
	}

	/// <summary>
	/// Parses the document for order-info.
	/// </summary>
	public static ExportOrder ParseOrder(string filePath) {
		using var wordDocument = WordprocessingDocument.Open(filePath, false);
		var coreProperties = wordDocument.PackageProperties ?? throw new NullReferenceException("The document has no core-properties!");
		var desc = coreProperties.Description;
		var descAsSpan = desc.AsSpan();
		if (string.IsNullOrWhiteSpace(desc))
			throw new InvalidDataException("The desc was null or white space. Unable to find order-id!");


		int startIndex = descAsSpan.IndexOf(MARKER_START);
		int endIndex = descAsSpan.IndexOf(MARKER_END);
		if (startIndex == -1) throw new InvalidDataContractException("The MARKER_START not found!");
		if (endIndex == -1) throw new InvalidDataContractException("The MARKER_END not found!");
		ReadOnlySpan<char> jsonContent = descAsSpan.Slice(startIndex + MARKER_START.Length, endIndex-startIndex-MARKER_START.Length);
		return ExportOrder.Parse(jsonContent) ?? throw new InvalidDataContractException("Malformed export-order json!");
	}

	/// <summary>
	/// Parses the given docuemnt for the order-id.
	/// </summary>
	public static Guid OrderId(string filePath) {
		using var wordDocument = WordprocessingDocument.Open(filePath, false);
		var coreProperties = wordDocument.PackageProperties ?? throw new NullReferenceException("The document has no core-properties!");
		var description = coreProperties.Description;
		if (string.IsNullOrWhiteSpace(description))
			throw new InvalidDataException("The desc was null or white space. Unable to find order-id!");

		var index = description.IndexOf(ORDER_ID_IDENTIFIER);
		if (index == -1) throw new InvalidDataException("order-id keyword not found!"); ;

		var startIndex = index + ORDER_ID_IDENTIFIER.Length;
		var orderIdSpan = description.AsSpan()[startIndex..(startIndex + 36)];
		return Guid.TryParse(orderIdSpan, out var orderId) ? orderId
			: throw new Exception($@"Invalid order id. Failed to parse ""{orderIdSpan}"" as a Guid.");
	}
}
