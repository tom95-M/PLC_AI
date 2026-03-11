using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using S7.Net;
using S7PlcDataTester.Wpf.Interfaces;
using S7PlcDataTester.Wpf.Models;

namespace S7PlcDataTester.Wpf.Communicators
{
    public class SiemensS7PlcCommunicator : IPlcCommunicator
    {
        private static readonly TimeSpan ConnectTimeout = TimeSpan.FromSeconds(5);
        private readonly Plc _plc;
        private readonly PlcConfig _config;
        
        public bool IsConnected => _plc?.IsConnected ?? false;
        
        public event EventHandler<PlcConnectionEventArgs>? ConnectionStateChanged;
        
        public SiemensS7PlcCommunicator(PlcConfig config)
        {
            _config = config;
            var plcType = CpuType.S71200;
            _plc = new Plc(plcType, config.IpAddress, short.Parse(config.Rack), short.Parse(config.Slot));
        }
        
        public async Task<bool> ConnectAsync()
        {
            try
            {
                await Task.Run(() => _plc.Open()).WaitAsync(ConnectTimeout);
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
                OnConnectionStateChanged(false, $"Operation failed: {ex.Message}");
                return false;
            }
        }

        public async Task DisconnectAsync()
        {
            try
            {
                await Task.Run(() => _plc.Close());
                OnConnectionStateChanged(false, "Disconnected");
            }
            catch (Exception ex)
            {
                OnConnectionStateChanged(false, $"Operation failed: {ex.Message}");
            }
        }

        public async Task<object> ReadAsync(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(address))
                {
                    throw new ArgumentException("Invalid argument.");
                }

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                // ?
                string standardAddress = ConvertToStandardAddress(address);
                object value = await Task.Run(() => _plc.Read(standardAddress));
                return value;
            }
            catch (ArgumentException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        public async Task<object> ReadAsync(string address, string dataType)
        {
            try
            {
                if (string.IsNullOrEmpty(address))
                {
                    throw new ArgumentException("Invalid argument.");
                }

                if (string.IsNullOrEmpty(dataType))
                {
                    throw new ArgumentException("Invalid argument.");
                }

                var normalizedDataType = NormalizeDataType(dataType);

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                // ?
                string standardAddress = ConvertToStandardAddress(address, normalizedDataType);
                
                // LREAL
                if (normalizedDataType == "lreal")
                {
                    // B
                    var (dbNumber, startByte) = ParseAddressForLReal(standardAddress);
                    // 8?
                    byte[] bytes = await Task.Run(() => _plc.ReadBytes(S7.Net.DataType.DataBlock, dbNumber, startByte, 8));
                    return SiemensBytesToDouble(bytes);
                }
                else if (normalizedDataType == "string")
                {
                    var stringLength = ParseStringLength(dataType, 20);
                    var (dbNumber, startByte) = ParseAddressForByteAccess(standardAddress);
                    if (stringLength == 0)
                    {
                        var headerBytes = await Task.Run(() => _plc.ReadBytes(S7.Net.DataType.DataBlock, dbNumber, startByte, 2));
                        var maxLen = headerBytes.Length > 0 ? headerBytes[0] : (byte)0;
                        var curLen = headerBytes.Length > 1 ? headerBytes[1] : (byte)0;
                        if (curLen > maxLen)
                        {
                            curLen = maxLen;
                        }

                        var textBytes = curLen > 0
                            ? await Task.Run(() => _plc.ReadBytes(S7.Net.DataType.DataBlock, dbNumber, startByte + 2, curLen))
                            : Array.Empty<byte>();
                        var text = Encoding.ASCII.GetString(textBytes);
                        return $"0x{maxLen:X2}0x{curLen:X2}{text}";
                    }

                    var bytes = await Task.Run(() => _plc.ReadBytes(S7.Net.DataType.DataBlock, dbNumber, startByte, stringLength + 2));

                    // S7 STRING format: [maxLen][curLen][chars...]
                    var maxLenFromData = bytes.Length > 0 ? bytes[0] : (byte)0;
                    var curLenFromData = bytes.Length > 1 ? bytes[1] : (byte)0;
                    if (maxLenFromData > 0 && curLenFromData <= maxLenFromData && bytes.Length >= 2 + curLenFromData)
                    {
                        return Encoding.ASCII.GetString(bytes, 2, curLenFromData);
                    }

                    // Backward compatibility for previous raw ASCII writes.
                    return Encoding.ASCII.GetString(bytes, 0, Math.Min(stringLength, bytes.Length)).TrimEnd('\0');
                }
                else
                {
                    // 
                    object value = await Task.Run(() => _plc.Read(standardAddress));
                    // ?
                    return ConvertToExpectedType(value, normalizedDataType);
                }
            }
            catch (ArgumentException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        public async Task WriteAsync(string address, object value)
        {
            await WriteAsync(address, value, "Bool");
        }

        public async Task WriteAsync(string address, object value, string dataType)
        {
            try
            {
                if (string.IsNullOrEmpty(address))
                {
                    throw new ArgumentException("Invalid argument.");
                }

                if (value == null)
                {
                    throw new ArgumentNullException("value", "Value cannot be null.");
                }

                if (string.IsNullOrEmpty(dataType))
                {
                    throw new ArgumentException("Invalid argument.");
                }

                var normalizedDataType = NormalizeDataType(dataType);

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                // ?
                string standardAddress = ConvertToStandardAddress(address, normalizedDataType);
                
                // LREAL
                if (normalizedDataType == "lreal")
                {
                    // B
                    var (dbNumber, startByte) = ParseAddressForLReal(standardAddress);
                    // double
                    double doubleValue = Convert.ToDouble(value);
                    // ?
                    byte[] bytes = DoubleToSiemensBytes(doubleValue);
                    await Task.Run(() => _plc.WriteBytes(S7.Net.DataType.DataBlock, dbNumber, startByte, bytes));
                }
                else if (normalizedDataType == "string")
                {
                    var stringLength = ParseStringLength(dataType, 20);
                    var (dbNumber, startByte) = ParseAddressForByteAccess(standardAddress);
                    var buffer = new byte[stringLength + 2];
                    var source = Encoding.ASCII.GetBytes(Convert.ToString(value) ?? string.Empty);
                    var actualLength = Math.Min(source.Length, stringLength);
                    buffer[0] = (byte)Math.Min(stringLength, 254);
                    buffer[1] = (byte)actualLength;
                    Array.Copy(source, 0, buffer, 2, actualLength);
                    await Task.Run(() => _plc.WriteBytes(S7.Net.DataType.DataBlock, dbNumber, startByte, buffer));
                }
                else
                {
                    // 
                    // ?
                    object convertedValue = ConvertToPlcType(value, normalizedDataType);
                    await Task.Run(() => _plc.Write(standardAddress, convertedValue));
                }
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (ArgumentException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        /// <summary>
        /// ?
        /// </summary>
        /// <param name="address">?</param>
        /// <returns>?</returns>
        private string ConvertToStandardAddress(string address)
        {
            return ConvertToStandardAddress(address, "Bool");
        }

        /// <summary>
        /// ?
        /// </summary>
        /// <param name="address">?</param>
        /// <param name="dataType"></param>
        /// <returns>?</returns>
        private string ConvertToStandardAddress(string address, string dataType)
        {
            if (string.IsNullOrEmpty(address))
            {
                throw new ArgumentException("Invalid argument.");
            }

            // ? (?DB2.DBX0.0 ?DB2.DBW2)
            if (address.Contains(".DBX") || address.Contains(".DBB") || address.Contains(".DBW") || address.Contains(".DBD"))
            {
                return address;
            }

            // ?DB2.0.0
            string[] parts = address.Split('.');
            if (parts.Length < 3)
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            string dbPart = parts[0];
            string byteOffset = parts[1];
            string bitOffset = parts[2];

            // DB
            int dbNumber;
            if (dbPart.StartsWith("DB", StringComparison.OrdinalIgnoreCase))
            {
                dbNumber = int.Parse(dbPart.Substring(2));
            }
            else
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            // ?
            string addressFormat = GetAddressFormatByDataType(dataType);
            
            // 
            if (addressFormat == "DBX")
            {
                return $"DB{dbNumber}.{addressFormat}{byteOffset}.{bitOffset}";
            }
            else
            {
                // OOL
                return $"DB{dbNumber}.{addressFormat}{byteOffset}";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        private string GetAddressFormatByDataType(string dataType)
        {
            switch (NormalizeDataType(dataType))
            {
                case "bool":
                    return "DBX";
                case "byte":
                    return "DBB";
                case "word":
                case "int":
                    return "DBW";
                case "dword":
                case "dint":
                case "real":
                    return "DBD";
                case "lreal":
                    // LREALBD??
                    return "DBD";
                case "string":
                    return "DBB";
                default:
                    return "DBX";
            }
        }

        /// <summary>
        /// LC??
        /// </summary>
        /// <param name="value">PLC??/param>
        /// <param name="dataType">?/param>
        /// <returns>?/returns>
        private object ConvertToExpectedType(object value, string dataType)
        {
            if (value == null)
            {
                return null;
            }

            switch (NormalizeDataType(dataType))
            {
                case "bool":
                    return Convert.ToBoolean(value);
                case "byte":
                    return Convert.ToByte(value);
                case "word":
                    return Convert.ToUInt16(value);
                case "int":
                    return Convert.ToInt16(value);
                case "dword":
                    return Convert.ToUInt32(value);
                case "dint":
                    return Convert.ToInt32(value);
                case "real":
                    // REAL7.Netloat??
                    double doubleValue;
                    if (value is int intValue1)
                    {
                        // ntyte
                        byte[] bytes1 = BitConverter.GetBytes(intValue1);
                        // loat
                        float floatValue1 = BitConverter.ToSingle(bytes1, 0);
                        // ouble
                        doubleValue = (double)floatValue1;
                    }
                    else if (value is uint uintValue1)
                    {
                        // intyte
                        byte[] bytes2 = BitConverter.GetBytes(uintValue1);
                        // loat
                        float floatValue2 = BitConverter.ToSingle(bytes2, 0);
                        // ouble
                        doubleValue = (double)floatValue2;
                    }
                    else if (value is long longValue)
                    {
                        // ongntbyte
                        int intValue2 = (int)longValue;
                        byte[] bytes3 = BitConverter.GetBytes(intValue2);
                        // loat
                        float floatValue3 = BitConverter.ToSingle(bytes3, 0);
                        // ouble
                        doubleValue = (double)floatValue3;
                    }
                    else if (value is ulong ulongValue)
                    {
                        // longintbyte
                        uint uintValue2 = (uint)ulongValue;
                        byte[] bytes4 = BitConverter.GetBytes(uintValue2);
                        // loat
                        float floatValue4 = BitConverter.ToSingle(bytes4, 0);
                        // ouble
                        doubleValue = (double)floatValue4;
                    }
                    else
                    {
                        // 
                        doubleValue = Convert.ToDouble(value);
                    }
                    
                    // 
                    // REAL2-3?
                    double roundedValue = doubleValue;
                    
                    // ??
                    double fractionalPart = Math.Abs(doubleValue - Math.Floor(doubleValue));
                    
                    if (fractionalPart >= 0.01)
                    {
                        // ?2?
                        roundedValue = Math.Round(doubleValue, 2);
                    }
                    else if (fractionalPart >= 0.001)
                    {
                        // ???
                        roundedValue = Math.Round(doubleValue, 3);
                    }
                    else
                    {
                        // ??
                        roundedValue = Math.Round(doubleValue, 0);
                    }
                    
                    return roundedValue;
                case "lreal":
                    // 
                    // LREAL?0?
                    double lrealValue = Convert.ToDouble(value);
                    return Math.Round(lrealValue, 10);
                case "string":
                    return value.ToString();
                default:
                    return value;
            }
        }

        /// <summary>
        /// ?PLC?
        /// </summary>
        /// <param name="value">?/param>
        /// <param name="dataType"></param>
        /// <returns>?/returns>
        private object ConvertToPlcType(object value, string dataType)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value", "Value cannot be null.");
            }

            switch (NormalizeDataType(dataType))
            {
                case "bool":
                    return Convert.ToBoolean(value);
                case "byte":
                    return Convert.ToByte(value);
                case "word":
                    return Convert.ToUInt16(value);
                case "int":
                    return Convert.ToInt16(value);
                case "dword":
                    return Convert.ToUInt32(value);
                case "dint":
                    return Convert.ToInt32(value);
                case "real":
                    // REALfloat
                    if (value is double doubleValue)
                    {
                        return (float)doubleValue;
                    }
                    return (float)Convert.ToDouble(value);
                case "lreal":
                    // LREALdouble
                    return Convert.ToDouble(value);
                case "string":
                    return value.ToString();
                default:
                    return value;
            }
        }

        /// <summary>
        /// LREAL?B
        /// </summary>
        /// <param name="address">?</param>
        /// <returns>DB</returns>
        private (int dbNumber, int startByte) ParseAddressForLReal(string address)
        {
            // DB1.DBD0
            string[] parts = address.Split('.');
            if (parts.Length < 2)
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            // DB?
            string dbPart = parts[0];
            int dbNumber;
            if (dbPart.StartsWith("DB", StringComparison.OrdinalIgnoreCase))
            {
                dbNumber = int.Parse(dbPart.Substring(2));
            }
            else
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            // 
            string dbdPart = parts[1];
            if (dbdPart.StartsWith("DBD", StringComparison.OrdinalIgnoreCase))
            {
                int startByte = int.Parse(dbdPart.Substring(3));
                return (dbNumber, startByte);
            }
            else
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }
        }

        private (int dbNumber, int startByte) ParseAddressForByteAccess(string address)
        {
            string[] parts = address.Split('.');
            if (parts.Length < 2)
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            if (!parts[0].StartsWith("DB", StringComparison.OrdinalIgnoreCase) ||
                !int.TryParse(parts[0].Substring(2), out var dbNumber))
            {
                throw new ArgumentException($"Invalid address format: {address}");
            }

            var offsetPart = parts[1];
            if (offsetPart.StartsWith("DBD", StringComparison.OrdinalIgnoreCase) ||
                offsetPart.StartsWith("DBW", StringComparison.OrdinalIgnoreCase) ||
                offsetPart.StartsWith("DBB", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(offsetPart.Substring(3), out var startByte))
                {
                    return (dbNumber, startByte);
                }
            }

            throw new ArgumentException($"Invalid address format: {address}");
        }

        private static double SiemensBytesToDouble(byte[] bytes)
        {
            if (bytes == null || bytes.Length != 8)
            {
                throw new ArgumentException($"LReal requires exactly 8 bytes, current: {bytes?.Length ?? 0}");
            }

            var buffer = (byte[])bytes.Clone();
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(buffer);
            }

            return BitConverter.ToDouble(buffer, 0);
        }

        private static byte[] DoubleToSiemensBytes(double value)
        {
            var bytes = BitConverter.GetBytes(value);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(bytes);
            }

            return bytes;
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

            if (int.TryParse(dataType[(separator + 1)..], out var length) && length >= 0)
            {
                return length;
            }

            return defaultLength;
        }

        public async Task<Dictionary<string, object>> ReadMultipleAsync(IEnumerable<string> addresses)
        {
            try
            {
                if (addresses == null)
                {
                    throw new ArgumentNullException("addresses", "Address list cannot be null.");
                }

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                var result = new Dictionary<string, object>();
                foreach (var address in addresses)
                {
                    result[address] = await ReadAsync(address);
                }
                return result;
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        /// <summary>
        /// ?
        /// </summary>
        /// <param name="addressDataTypePairs">?</param>
        /// <returns>?</returns>
        public async Task<Dictionary<string, object>> ReadMultipleAsync(Dictionary<string, string> addressDataTypePairs)
        {
            try
            {
                if (addressDataTypePairs == null)
                {
                    throw new ArgumentNullException("addressDataTypePairs", "Address and data type pairs cannot be null.");
                }

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                var result = new Dictionary<string, object>();
                foreach (var kvp in addressDataTypePairs)
                {
                    result[kvp.Key] = await ReadAsync(kvp.Key, kvp.Value);
                }
                return result;
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        public async Task WriteMultipleAsync(Dictionary<string, object> values)
        {
            try
            {
                if (values == null)
                {
                    throw new ArgumentNullException("values", "Values cannot be null.");
                }

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                foreach (var kvp in values)
                {
                    await WriteAsync(kvp.Key, kvp.Value);
                }
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="valuesWithDataType">?</param>
        public async Task WriteMultipleAsync(IEnumerable<(string address, object value, string dataType)> valuesWithDataType)
        {
            try
            {
                if (valuesWithDataType == null)
                {
                    throw new ArgumentNullException("valuesWithDataType", "Values with data type cannot be null.");
                }

                if (!IsConnected)
                {
                    throw new InvalidOperationException("PLC is not connected.");
                }

                foreach (var (address, value, dataType) in valuesWithDataType)
                {
                    await WriteAsync(address, value, dataType);
                }
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Operation failed: {ex.Message}");
            }
        }


        private static string NormalizeDataType(string dataType)
        {
            if (string.IsNullOrWhiteSpace(dataType))
            {
                return string.Empty;
            }

            var separator = dataType.IndexOf(':');
            var normalized = separator >= 0 ? dataType[..separator] : dataType;
            return normalized.Trim().ToLowerInvariant();
        }
        private void OnConnectionStateChanged(bool isConnected, string message)
        {
            ConnectionStateChanged?.Invoke(this, new PlcConnectionEventArgs(isConnected, message));
        }
    }
}
