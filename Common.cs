global using static WordHandler.Common;
global using static Common.Ancillary.Common;
global using static Common.Ancillary.Utilities;

namespace WordHandler;
internal static class Common {
	public const string MARKER_START = "\n#_RAW_START_#\n";
	public const string MARKER_END = "\n#_RAW_END_#\n";

	public const string ORDER_ID_IDENTIFIER = "order-id: ";
}
