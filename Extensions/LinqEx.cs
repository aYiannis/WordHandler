using DocumentFormat.OpenXml;

namespace WordHandler.Extensions;

internal static class LinqEx {
	public static int FindIndex(this OpenXmlElementList self, OpenXmlElement element) {
		var count = self.Count;
		for (var i = 0; i < count; i++) {
			if (self[i] == element)
				return i;
		}
		return -1;
	}
}
