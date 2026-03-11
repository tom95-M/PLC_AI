using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace S7PlcDataTester.Wpf.Interfaces
{
    public interface IPlcCommunicator
    {
        Task<bool> ConnectAsync();
        Task DisconnectAsync();
        bool IsConnected { get; }
        
        Task<object> ReadAsync(string address);
        Task<object> ReadAsync(string address, string dataType);
        Task WriteAsync(string address, object value);
        Task WriteAsync(string address, object value, string dataType);
        Task<Dictionary<string, object>> ReadMultipleAsync(IEnumerable<string> addresses);
        Task<Dictionary<string, object>> ReadMultipleAsync(Dictionary<string, string> addressDataTypePairs);
        Task WriteMultipleAsync(Dictionary<string, object> values);
        Task WriteMultipleAsync(IEnumerable<(string address, object value, string dataType)> valuesWithDataType);
        
        event EventHandler<PlcConnectionEventArgs> ConnectionStateChanged;
    }
    
    public class PlcConnectionEventArgs : EventArgs
    {
        public bool IsConnected { get; }
        public string? Message { get; }
        
        public PlcConnectionEventArgs(bool isConnected, string? message = null)
        {
            IsConnected = isConnected;
            Message = message;
        }
    }
}
