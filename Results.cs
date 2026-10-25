using System.Collections;

namespace WinFormsApp1
{
    public sealed class Results : IReadOnlyDictionary<string, WorkStatus>
    {
        private readonly Dictionary<string, WorkStatus> _items;

        public Results() : this([]) { }
        public Results(Dictionary<string, WorkStatus> items) => _items = items ?? [];

        public int Ratio => _items.Count(kv => kv.Value == WorkStatus.Зачтено);

        public Results Filter(string pattern)
        {
            return new(
                _items
                    .Where(kv => kv.Key.Contains(pattern, StringComparison.Ordinal))
                    .ToDictionary(kv => kv.Key, kv => kv.Value)
                    );
        }

        public int Count => _items.Count;
        public IEnumerable<string> Keys => _items.Keys;
        public IEnumerable<WorkStatus> Values => _items.Values;
        public WorkStatus this[string key] => _items[key];
        public bool ContainsKey(string key) => _items.ContainsKey(key);
        public bool TryGetValue(string key, out WorkStatus value) => _items.TryGetValue(key, out value);
        public IEnumerator<KeyValuePair<string, WorkStatus>> GetEnumerator() => _items.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public void Set(string key, WorkStatus status)
        {
            _items[key] = status;
        }
    }
}
