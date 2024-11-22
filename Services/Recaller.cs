using Common;

using Model.Entities;

namespace WordHandler.Services;

public class Recaller {
	static Recaller? _current = null;
	public static Recaller Current => _current is null ? throw new NullReferenceException() : _current;

	readonly Dictionary<Guid, ExportOrder> orders = [];
	public ExportOrder[] Orders => [.. orders.Values];

	public static Recaller FromFile(string? filePath = null) {
		filePath ??= Paths.History;

		Recaller recaller = new();
		if (File.Exists(filePath))
			recaller.Load(filePath);

		return recaller;
	}

	public static void Initialize() => Initialize(Paths.History);
	public static void Initialize(string filePath) {
		if (_current is not null) return;

		_current = FromFile(filePath);
	}

	public ExportOrder? ById(Guid id) => orders.TryGetValue(id, out var o) ? o : null;

	bool changed = false;
	public void Register(ExportOrder order) {
		changed = true;
		orders[order.Id] = order;
	}

	/// <summary> Populates the values based on the binary file. </summary>
	/// <param name="fpath"> The path to the binary file. </param>
	public void Load(string? fpath = null) {
		using var fileStream = File.OpenRead(fpath ?? Paths.History);
		using var binaryReader = new BinaryReader(fileStream);

		var count = binaryReader.ReadInt32();
		orders.Clear();
		for (var i = 0; i < count; i++) {
			var order = ExportOrder.ReadFrom(binaryReader);
			orders.Add(order.Id, order);
		}
	}

	/// <summary> Stores all the orders into the history.bin file. </summary>
	public void Save() {
		if (!changed) return;

		using var fileStream = File.OpenWrite(Paths.History);
		using BinaryWriter binaryWriter = new(fileStream);
		binaryWriter.Write(orders.Count);
		foreach (var order in orders.Values)
			order.WriteTo(binaryWriter);
		changed = false;
	}
}
