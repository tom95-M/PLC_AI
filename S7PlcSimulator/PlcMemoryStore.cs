using Snap7;

namespace S7PlcSimulator;

public sealed class PlcMemoryStore
{
    private readonly Dictionary<int, byte[]> _dbBlocks;
    private readonly object _sync = new();

    public PlcMemoryStore(Dictionary<int, int> dbSizes)
    {
        if (dbSizes.Count == 0)
        {
            throw new ArgumentException("At least one DB block is required.", nameof(dbSizes));
        }

        _dbBlocks = dbSizes.ToDictionary(x => x.Key, x => new byte[x.Value]);
    }

    public IReadOnlyDictionary<int, byte[]> DbBlocks => _dbBlocks;

    public IReadOnlyDictionary<int, byte[]> GetDbBlocksSnapshot()
    {
        lock (_sync)
        {
            return _dbBlocks.ToDictionary(x => x.Key, x => x.Value);
        }
    }

    public byte[] EnsureDb(int db, int size, out bool created)
    {
        if (db <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(db), "DB number must be greater than 0.");
        }

        if (size <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(size), "DB size must be greater than 0.");
        }

        lock (_sync)
        {
            if (_dbBlocks.TryGetValue(db, out var existing))
            {
                created = false;
                return existing;
            }

            var block = new byte[size];
            _dbBlocks[db] = block;
            created = true;
            return block;
        }
    }

    public bool TryGetDb(int db, out byte[]? block)
    {
        lock (_sync)
        {
            return _dbBlocks.TryGetValue(db, out block);
        }
    }

    public byte[] Read(int db, int start, int size)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, size);

            var output = new byte[size];
            Buffer.BlockCopy(block, start, output, 0, size);
            return output;
        }
    }

    public void Write(int db, int start, byte[] data)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, data.Length);
            Buffer.BlockCopy(data, 0, block, start, data.Length);
        }
    }

    public bool GetBool(int db, int byteIndex, int bitIndex)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, byteIndex, 1);
            return S7.GetBitAt(block, byteIndex, bitIndex);
        }
    }

    public void SetBool(int db, int byteIndex, int bitIndex, bool value)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, byteIndex, 1);
            S7.SetBitAt(ref block, byteIndex, bitIndex, value);
        }
    }

    public short GetInt(int db, int start)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 2);
            return (short)S7.GetIntAt(block, start);
        }
    }

    public void SetInt(int db, int start, short value)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 2);
            S7.SetIntAt(block, start, value);
        }
    }

    public int GetDInt(int db, int start)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 4);
            return S7.GetDIntAt(block, start);
        }
    }

    public void SetDInt(int db, int start, int value)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 4);
            S7.SetDIntAt(block, start, value);
        }
    }

    public float GetReal(int db, int start)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 4);
            return S7.GetRealAt(block, start);
        }
    }

    public void SetReal(int db, int start, float value)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 4);
            S7.SetRealAt(block, start, value);
        }
    }

    public string GetString(int db, int start)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, 2);
            var maxLength = block[start];
            CheckRange(block, start, maxLength + 2);
            return S7.GetStringAt(block, start);
        }
    }

    public void SetString(int db, int start, int maxLength, string value)
    {
        lock (_sync)
        {
            var block = GetDb(db);
            CheckRange(block, start, maxLength + 2);
            S7.SetStringAt(block, start, maxLength, value);
        }
    }

    private byte[] GetDb(int db)
    {
        if (!_dbBlocks.TryGetValue(db, out var block))
        {
            throw new ArgumentOutOfRangeException(nameof(db), $"DB{db} not configured.");
        }

        return block;
    }

    private static void CheckRange(byte[] block, int start, int size)
    {
        if (start < 0 || size < 0 || start + size > block.Length)
        {
            throw new ArgumentOutOfRangeException($"Invalid range start={start}, size={size}, length={block.Length}.");
        }
    }
}
