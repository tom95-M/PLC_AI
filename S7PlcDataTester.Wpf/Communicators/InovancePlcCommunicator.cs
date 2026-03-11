using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Sockets;
using System.Threading.Tasks;
using NModbus;
using S7PlcDataTester.Wpf.Interfaces;
using S7PlcDataTester.Wpf.Models;

namespace S7PlcDataTester.Wpf.Communicators
{
    public class InovancePlcCommunicator : IPlcCommunicator
    {
        private enum MemoryArea
        {
            HoldingRegister,
            Coil
        }

        private static readonly TimeSpan ConnectTimeout = TimeSpan.FromSeconds(5);
        private TcpClient? _client;
        private IModbusMaster? _master;
        private readonly PlcConfig _config;

        public bool IsConnected => _client?.Connected ?? false;

        public event EventHandler<PlcConnectionEventArgs>? ConnectionStateChanged;

        public InovancePlcCommunicator(PlcConfig config)
        {
            _config = config;
            if (_config.Port == 102)
            {
                _config.Port = 502;
            }
        }

        public async Task<bool> ConnectAsync()
        {
            try
            {
                _client = new TcpClient
                {
                    ReceiveTimeout = (int)ConnectTimeout.TotalMilliseconds,
                    SendTimeout = (int)ConnectTimeout.TotalMilliseconds
                };

                await _client.ConnectAsync(_config.IpAddress, _config.Port).WaitAsync(ConnectTimeout);
                var factory = new ModbusFactory();
                _master = factory.CreateMaster(_client);
                OnConnectionStateChanged(true, "Connected");
                return true;
            }
            catch (TimeoutException)
            {
                OnConnectionStateChanged(false, $"Connect timeout ({(int)ConnectTimeout.TotalSeconds}s)");
                return false;
            }
            catch (Exception ex)
            {
                OnConnectionStateChanged(false, $"Connect failed: {ex.Message}");
                return false;
            }
        }

        public Task DisconnectAsync()
        {
            try
            {
                _master?.Dispose();
                _master = null;
                _client?.Close();
                _client = null;
                OnConnectionStateChanged(false, "Disconnected");
            }
            catch (Exception ex)
            {
                OnConnectionStateChanged(false, $"Disconnect failed: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        public async Task<object> ReadAsync(string address)
        {
            return await ReadAsync(address, "Word");
        }

        public async Task<object> ReadAsync(string address, string dataType)
        {
            try
            {
                EnsureConnected();
                var (area, startAddress) = ParseAddress(address);
                var normalizedType = NormalizeDataType(dataType);
                var stringLength = ParseStringLength(dataType, 20);
                var unitId = GetUnitId();

                if (area == MemoryArea.Coil)
                {
                    if (normalizedType != "bool")
                    {
                        throw new NotSupportedException("Coil address only supports Bool.");
                    }

                    var values = await Task.Run(() => _master!.ReadCoils(unitId, startAddress, 1));
                    return values.Length > 0 && values[0];
                }

                var wordCount = GetWordCount(normalizedType, stringLength);
                var registers = await Task.Run(() => _master!.ReadHoldingRegisters(unitId, startAddress, (ushort)wordCount));
                return ConvertRegistersToValue(registers, normalizedType, stringLength);
            }
            catch (Exception ex)
            {
                throw new Exception($"Read failed: {ex.Message}");
            }
        }

        public Task WriteAsync(string address, object value) => WriteAsync(address, value, "Bool");

        public async Task WriteAsync(string address, object value, string dataType)
        {
            try
            {
                EnsureConnected();
                var (area, startAddress) = ParseAddress(address);
                var normalizedType = NormalizeDataType(dataType);
                var stringLength = ParseStringLength(dataType, 20);
                var unitId = GetUnitId();

                if (area == MemoryArea.Coil)
                {
                    if (normalizedType != "bool")
                    {
                        throw new NotSupportedException("Coil address only supports Bool.");
                    }

                    var boolValue = Convert.ToBoolean(value, CultureInfo.InvariantCulture);
                    await Task.Run(() => _master!.WriteSingleCoil(unitId, startAddress, boolValue));
                    return;
                }

                var registers = ConvertValueToRegisters(value, normalizedType, stringLength);
                if (registers.Length == 1)
                {
                    await Task.Run(() => _master!.WriteSingleRegister(unitId, startAddress, registers[0]));
                }
                else
                {
                    await Task.Run(() => _master!.WriteMultipleRegisters(unitId, startAddress, registers));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Write failed: {ex.Message}");
            }
        }

        public async Task<Dictionary<string, object>> ReadMultipleAsync(IEnumerable<string> addresses)
        {
            try
            {
                var result = new Dictionary<string, object>();
                foreach (var address in addresses)
                {
                    result[address] = await ReadAsync(address);
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Batch read failed: {ex.Message}");
            }
        }

        public async Task WriteMultipleAsync(Dictionary<string, object> values)
        {
            try
            {
                foreach (var kvp in values)
                {
                    await WriteAsync(kvp.Key, kvp.Value);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Batch write failed: {ex.Message}");
            }
        }

        public async Task<Dictionary<string, object>> ReadMultipleAsync(Dictionary<string, string> addressDataTypePairs)
        {
            try
            {
                var result = new Dictionary<string, object>();
                foreach (var kvp in addressDataTypePairs)
                {
                    result[kvp.Key] = await ReadAsync(kvp.Key, kvp.Value);
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Batch read failed: {ex.Message}");
            }
        }

        public async Task WriteMultipleAsync(IEnumerable<(string address, object value, string dataType)> valuesWithDataType)
        {
            try
            {
                foreach (var (address, value, dataType) in valuesWithDataType)
                {
                    await WriteAsync(address, value, dataType);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Batch write failed: {ex.Message}");
            }
        }

        private void OnConnectionStateChanged(bool isConnected, string message)
        {
            ConnectionStateChanged?.Invoke(this, new PlcConnectionEventArgs(isConnected, message));
        }

        private void EnsureConnected()
        {
            if (_master == null || _client?.Connected != true)
            {
                throw new InvalidOperationException("PLC is not connected.");
            }
        }

        private static string NormalizeDataType(string dataType)
        {
            var type = dataType ?? string.Empty;
            var separator = type.IndexOf(':');
            if (separator >= 0)
            {
                type = type[..separator];
            }

            return type.Trim().ToLowerInvariant();
        }

        private static int ParseStringLength(string dataType, int defaultLength)
        {
            if (string.IsNullOrWhiteSpace(dataType))
            {
                return defaultLength;
            }

            var separator = dataType.IndexOf(':');
            if (separator < 0)
            {
                return defaultLength;
            }

            return int.TryParse(dataType[(separator + 1)..], out var length) && length > 0
                ? length
                : defaultLength;
        }

        private static int GetWordCount(string normalizedType, int stringLength)
        {
            return normalizedType switch
            {
                "bool" => 1,
                "byte" => 1,
                "word" => 1,
                "int" => 1,
                "dword" => 2,
                "dint" => 2,
                "real" => 2,
                "lreal" => 4,
                "string" => Math.Max(1, (stringLength + 1) / 2),
                _ => throw new NotSupportedException($"Unsupported data type: {normalizedType}")
            };
        }

        private static object ConvertRegistersToValue(ushort[] registers, string normalizedType, int stringLength)
        {
            return normalizedType switch
            {
                "bool" => registers[0] != 0,
                "byte" => (byte)(registers[0] & 0x00FF),
                "word" => registers[0],
                "int" => unchecked((short)registers[0]),
                "dword" => ((uint)registers[1] << 16) | registers[0],
                "dint" => unchecked((int)(((uint)registers[1] << 16) | registers[0])),
                "real" => BitConverter.ToSingle(new[]
                {
                    (byte)(registers[0] & 0xFF),
                    (byte)(registers[0] >> 8),
                    (byte)(registers[1] & 0xFF),
                    (byte)(registers[1] >> 8)
                }, 0),
                "lreal" => BitConverter.ToDouble(new[]
                {
                    (byte)(registers[0] & 0xFF),
                    (byte)(registers[0] >> 8),
                    (byte)(registers[1] & 0xFF),
                    (byte)(registers[1] >> 8),
                    (byte)(registers[2] & 0xFF),
                    (byte)(registers[2] >> 8),
                    (byte)(registers[3] & 0xFF),
                    (byte)(registers[3] >> 8)
                }, 0),
                "string" => DecodeString(registers, stringLength),
                _ => throw new NotSupportedException($"Unsupported data type: {normalizedType}")
            };
        }

        private static ushort[] ConvertValueToRegisters(object value, string normalizedType, int stringLength)
        {
            return normalizedType switch
            {
                "bool" => new[] { (ushort)(Convert.ToBoolean(value, CultureInfo.InvariantCulture) ? 1 : 0) },
                "byte" => new[] { (ushort)Convert.ToByte(value, CultureInfo.InvariantCulture) },
                "word" => new[] { Convert.ToUInt16(value, CultureInfo.InvariantCulture) },
                "int" => new[] { unchecked((ushort)Convert.ToInt16(value, CultureInfo.InvariantCulture)) },
                "dword" => UInt32ToWords(Convert.ToUInt32(value, CultureInfo.InvariantCulture)),
                "dint" => UInt32ToWords(unchecked((uint)Convert.ToInt32(value, CultureInfo.InvariantCulture))),
                "real" => BytesToWords(BitConverter.GetBytes(Convert.ToSingle(value, CultureInfo.InvariantCulture))),
                "lreal" => BytesToWords(BitConverter.GetBytes(Convert.ToDouble(value, CultureInfo.InvariantCulture))),
                "string" => EncodeString(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, stringLength),
                _ => throw new NotSupportedException($"Unsupported data type: {normalizedType}")
            };
        }

        private static ushort[] UInt32ToWords(uint value)
        {
            return new[] { (ushort)(value & 0xFFFF), (ushort)((value >> 16) & 0xFFFF) };
        }

        private static ushort[] BytesToWords(byte[] bytes)
        {
            var wordCount = (bytes.Length + 1) / 2;
            var words = new ushort[wordCount];
            for (var i = 0; i < wordCount; i++)
            {
                var low = bytes[i * 2];
                var high = (i * 2 + 1) < bytes.Length ? bytes[i * 2 + 1] : (byte)0;
                words[i] = (ushort)(low | (high << 8));
            }

            return words;
        }

        private static string DecodeString(ushort[] registers, int length)
        {
            var bytes = new byte[registers.Length * 2];
            for (var i = 0; i < registers.Length; i++)
            {
                bytes[i * 2] = (byte)(registers[i] & 0xFF);
                bytes[i * 2 + 1] = (byte)(registers[i] >> 8);
            }

            var raw = System.Text.Encoding.ASCII.GetString(bytes, 0, Math.Min(length, bytes.Length));
            return raw.TrimEnd('\0');
        }

        private static ushort[] EncodeString(string value, int length)
        {
            var bytes = new byte[Math.Max(1, length)];
            var src = System.Text.Encoding.ASCII.GetBytes(value);
            Array.Copy(src, bytes, Math.Min(bytes.Length, src.Length));
            return BytesToWords(bytes);
        }

        private static (MemoryArea area, ushort startAddress) ParseAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                throw new ArgumentException("Address is required.");
            }

            var normalized = address.Trim().ToUpperInvariant();
            if (normalized.StartsWith("D") && ushort.TryParse(normalized[1..], out var dAddress))
            {
                return (MemoryArea.HoldingRegister, dAddress);
            }

            if (normalized.StartsWith("M") && ushort.TryParse(normalized[1..], out var mAddress))
            {
                return (MemoryArea.Coil, mAddress);
            }

            if (ushort.TryParse(normalized, out var directAddress))
            {
                return (MemoryArea.HoldingRegister, directAddress);
            }

            throw new NotSupportedException($"Unsupported address format: {address}");
        }

        private byte GetUnitId()
        {
            return byte.TryParse(_config.Station, out var station) ? station : (byte)1;
        }
    }
}
