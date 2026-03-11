using Snap7;

namespace S7PlcSimulator;

public sealed class S7ServerHost : IDisposable
{
    private readonly PlcMemoryStore _memoryStore;
    private readonly S7Server _server = new();
    private readonly S7Server.TSrvCallback _eventCallback;
    private bool _started;
    private readonly HashSet<int> _registeredDbNumbers = [];
    private long _lastRwSummaryTick;
    private int _readEvents;
    private int _writeEvents;
    private readonly Action<string>? _log;

    public S7ServerHost(PlcMemoryStore memoryStore, Action<string>? log = null)
    {
        _memoryStore = memoryStore;
        _eventCallback = OnServerEvent;
        _log = log;
    }

    public void Start(string localAddress, int port)
    {
        if (_started)
        {
            return;
        }

        RegisterDbAreasIfNeeded();
        _server.EventMask = S7Server.evcAll;
        _server.SetEventsCallBack(_eventCallback, IntPtr.Zero);

        short serverPort = checked((short)port);
        var setPortResult = _server.SetParam(S7Consts.p_u16_LocalPort, ref serverPort);
        if (setPortResult != 0)
        {
            throw new InvalidOperationException($"SetParam(LocalPort) failed: {_server.ErrorText(setPortResult)}");
        }

        var startResult = _server.StartTo(localAddress);
        if (startResult != 0)
        {
            throw new InvalidOperationException($"StartTo({localAddress}:{port}) failed: {_server.ErrorText(startResult)}");
        }

        _started = true;
        WriteLog($"[SERVER] Started at {localAddress}:{port}");
    }

    public void Stop()
    {
        if (!_started)
        {
            return;
        }

        var stopResult = _server.Stop();
        if (stopResult != 0)
        {
            WriteLog($"[SERVER] Stop failed: {_server.ErrorText(stopResult)}");
        }
        else
        {
            WriteLog("[SERVER] Stopped");
        }

        _started = false;
    }

    public void Dispose()
    {
        Stop();
    }

    public void EnsureDbAreaRegistered(int dbNumber)
    {
        if (!_memoryStore.TryGetDb(dbNumber, out var block) || block is null)
        {
            throw new InvalidOperationException($"DB{dbNumber} not found in memory store.");
        }

        RegisterDbArea(dbNumber, block);
    }

    private void RegisterDbAreasIfNeeded()
    {
        foreach (var block in _memoryStore.GetDbBlocksSnapshot())
        {
            RegisterDbArea(block.Key, block.Value);
        }
    }

    private void RegisterDbArea(int dbNumber, byte[] dbData)
    {
        if (_registeredDbNumbers.Contains(dbNumber))
        {
            return;
        }

        var rc = _server.RegisterArea(S7Server.srvAreaDB, dbNumber, ref dbData, dbData.Length);
        if (rc != 0)
        {
            throw new InvalidOperationException($"Register DB{dbNumber} failed: {_server.ErrorText(rc)}");
        }

        _registeredDbNumbers.Add(dbNumber);
        WriteLog($"[SERVER] Registered DB{dbNumber}, size={dbData.Length}");
    }

    private void OnServerEvent(IntPtr usrPtr, ref S7Server.USrvEvent evt, int size)
    {
        var text = _server.EventText(ref evt);
        WriteLog($"[EVENT] {text}");
        if (text.Contains("Read request", StringComparison.OrdinalIgnoreCase))
        {
            Interlocked.Increment(ref _readEvents);
        }
        else if (text.Contains("Write request", StringComparison.OrdinalIgnoreCase))
        {
            Interlocked.Increment(ref _writeEvents);
        }

        var nowTick = Environment.TickCount64;
        var previousTick = Interlocked.Read(ref _lastRwSummaryTick);
        if (nowTick - previousTick < 1000 ||
            Interlocked.CompareExchange(ref _lastRwSummaryTick, nowTick, previousTick) != previousTick)
        {
            return;
        }

        var readCount = Interlocked.Exchange(ref _readEvents, 0);
        var writeCount = Interlocked.Exchange(ref _writeEvents, 0);
        if (readCount > 0 || writeCount > 0)
        {
            WriteLog($"[RW] 1s summary: READ={readCount}, WRITE={writeCount}");
        }
    }

    private void WriteLog(string message)
    {
        _log?.Invoke(message);
    }
}
