using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Sockets;
using System.Threading.Tasks;
using S7PlcDataTester.Wpf.Interfaces;
using S7PlcDataTester.Wpf.Models;

namespace S7PlcDataTester.Wpf.Communicators
{
    public class KeyenceKvPlcCommunicator : IPlcCommunicator
    {
        private enum DeviceArea
        {
            Word,
            Bit
        }

        private static readonly TimeSpan ConnectTimeout = TimeSpan.FromSeconds(5);
        private const byte DeviceCodeD = 0xA8;
        private const byte DeviceCodeM = 0x90;

        private TcpClient? _client;
        private NetworkStream? _stream;
        private readonly PlcConfig _config;

        public bool IsConnected => _client?.Connected ?? false;

        public event EventHandler<PlcConnectionEventArgs>? ConnectionStateChanged;

        public KeyenceKvPlcCommunicator(PlcConfig config)
        {
            _config = config;
            if (_config.Port == 102)
            {
                _config.Port = 5000;
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
                _stream = _client.GetStream();
                _stream.ReadTimeout = (int)ConnectTimeout.TotalMilliseconds;
                _stream.WriteTimeout = (int)ConnectTimeout.TotalMilliseconds;
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
                _stream?.Close();
                _stream = null;
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

        public Task<object> ReadAsync(string address) => ReadAsync(address, "Word");

        public async Task<object> ReadAsync(string address, string dataType)
        {
            try
            {
                EnsureConnected();
                var (area, startAddress) = ParseAddress(address);
                var normalizedType = NormalizeDataType(dataType);
                var stringLength = ParseStringLength(dataType, 20);

                if (area == DeviceArea.Bit)
                {
                    if (normalizedType != "bool")
                    {
                        throw new NotSupportedException("Bit addresses (M/MR) only support Bool.");
                    }

                    return await ReadBitAsync(startAddress);
                }

                var wordCount = GetWordCount(normalizedType, stringLength);
                var words = await ReadWordsAsync(startAddress, wordCount);
                return ConvertRegistersToValue(words, normalizedType, stringLength);
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

                if (area == DeviceArea.Bit)
                {
                    if (normalizedType != "bool")
                    {
                        throw new NotSupportedException("Bit addresses (M/MR) only support Bool.");
                    }

                    var boolValue = Convert.ToBoolean(value, CultureInfo.InvariantCulture);
                    await WriteBitAsync(startAddress, boolValue);
                    return;
                }

                var words = ConvertValueToRegisters(value, normalizedType, stringLength);
                await WriteWordsAsync(startAddress, words);
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
            if (_stream == null || _client?.Connected != true)
            {
                throw new InvalidOperationException("PLC is not connected.");
            }
        }

        private static (DeviceArea area, int startAddress) ParseAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                throw new ArgumentException("Address is required.");
            }

            var normalized = address.Trim().ToUpperInvariant();
            if ((normalized.StartsWith("DM") && int.TryParse(normalized[2..], out var dmAddress) && dmAddress >= 0) ||
                (normalized.StartsWith("D") && int.TryParse(normalized[1..], out dmAddress) && dmAddress >= 0))
            {
                return (DeviceArea.Word, dmAddress);
            }

            if ((normalized.StartsWith("MR") && int.TryParse(normalized[2..], out var mrAddress) && mrAddress >= 0) ||
                (normalized.StartsWith("M") && int.TryParse(normalized[1..], out mrAddress) && mrAddress >= 0))
            {
                return (DeviceArea.Bit, mrAddress);
            }

            throw new NotSupportedException($"Unsupported KV address format: {address}. Examples: DM100, D100, MR10, M10");
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

        private async Task<ushort[]> ReadWordsAsync(int startAddress, int points)
        {
            var body = new List<byte>
            {
                (byte)(startAddress & 0xFF),
                (byte)((startAddress >> 8) & 0xFF),
                (byte)((startAddress >> 16) & 0xFF),
                DeviceCodeD,
                (byte)(points & 0xFF),
                (byte)((points >> 8) & 0xFF)
            };

            var data = await SendMcRequestAsync(0x0401, 0x0000, body.ToArray());
            if (data.Length < points * 2)
            {
                throw new IOException("PLC response data length mismatch.");
            }

            var result = new ushort[points];
            for (var i = 0; i < points; i++)
            {
                result[i] = (ushort)(data[i * 2] | (data[i * 2 + 1] << 8));
            }

            return result;
        }

        private async Task<bool> ReadBitAsync(int startAddress)
        {
            var body = new List<byte>
            {
                (byte)(startAddress & 0xFF),
                (byte)((startAddress >> 8) & 0xFF),
                (byte)((startAddress >> 16) & 0xFF),
                DeviceCodeM,
                0x01,
                0x00
            };

            var data = await SendMcRequestAsync(0x0401, 0x0001, body.ToArray());
            if (data.Length < 1)
            {
                throw new IOException("PLC response data length mismatch.");
            }

            return (data[0] & 0x01) != 0;
        }

        private async Task WriteWordsAsync(int startAddress, ushort[] words)
        {
            var body = new List<byte>
            {
                (byte)(startAddress & 0xFF),
                (byte)((startAddress >> 8) & 0xFF),
                (byte)((startAddress >> 16) & 0xFF),
                DeviceCodeD,
                (byte)(words.Length & 0xFF),
                (byte)((words.Length >> 8) & 0xFF)
            };

            foreach (var word in words)
            {
                body.Add((byte)(word & 0xFF));
                body.Add((byte)(word >> 8));
            }

            _ = await SendMcRequestAsync(0x1401, 0x0000, body.ToArray());
        }

        private async Task WriteBitAsync(int startAddress, bool value)
        {
            var body = new List<byte>
            {
                (byte)(startAddress & 0xFF),
                (byte)((startAddress >> 8) & 0xFF),
                (byte)((startAddress >> 16) & 0xFF),
                DeviceCodeM,
                0x01,
                0x00,
                (byte)(value ? 0x01 : 0x00)
            };

            _ = await SendMcRequestAsync(0x1401, 0x0001, body.ToArray());
        }

        private async Task<byte[]> SendMcRequestAsync(ushort command, ushort subCommand, byte[] body)
        {
            if (_stream == null)
            {
                throw new InvalidOperationException("PLC stream is not available.");
            }

            var frame = BuildMc3EFrame(command, subCommand, body);
            await _stream.WriteAsync(frame, 0, frame.Length);
            await _stream.FlushAsync();

            var header = await ReadExactAsync(_stream, 9);
            var responseLength = header[7] | (header[8] << 8);
            var payload = await ReadExactAsync(_stream, responseLength);
            if (payload.Length < 2)
            {
                throw new IOException("PLC response payload is invalid.");
            }

            var endCode = (ushort)(payload[0] | (payload[1] << 8));
            if (endCode != 0)
            {
                throw new IOException($"MC protocol error code: 0x{endCode:X4}");
            }

            var data = new byte[payload.Length - 2];
            Buffer.BlockCopy(payload, 2, data, 0, data.Length);
            return data;
        }

        private static byte[] BuildMc3EFrame(ushort command, ushort subCommand, byte[] body)
        {
            var requestLength = 2 + 2 + 2 + body.Length;
            var frame = new List<byte>
            {
                0x50,
                0x00,
                0x00,
                0xFF,
                0xFF,
                0x03,
                0x00,
                (byte)(requestLength & 0xFF),
                (byte)((requestLength >> 8) & 0xFF),
                0x10,
                0x00,
                (byte)(command & 0xFF),
                (byte)((command >> 8) & 0xFF),
                (byte)(subCommand & 0xFF),
                (byte)((subCommand >> 8) & 0xFF)
            };

            frame.AddRange(body);
            return frame.ToArray();
        }

        private static async Task<byte[]> ReadExactAsync(NetworkStream stream, int length)
        {
            var buffer = new byte[length];
            var offset = 0;
            while (offset < length)
            {
                var read = await stream.ReadAsync(buffer, offset, length - offset);
                if (read <= 0)
                {
                    throw new IOException("PLC connection closed unexpectedly.");
                }

                offset += read;
            }

            return buffer;
        }
    }
}
