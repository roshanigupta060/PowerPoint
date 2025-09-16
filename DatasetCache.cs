using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptExcelSync
{
    public static class DatasetCache
    {
        private static readonly Dictionary<string, DataTable> _cache = new Dictionary<string, DataTable>();

        public static DataTable GetOrLoad(string path)
        {
            if (_cache.ContainsKey(path))
                return _cache[path];

            var dt = new DatasetManager().LoadDataset(path);
            _cache[path] = dt;
            return dt;
        }

        public static void Clear(string path = null)
        {
            if (path == null)
                _cache.Clear();
            else if (_cache.ContainsKey(path))
                _cache.Remove(path);
        }
    }

}
