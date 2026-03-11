using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using S7PlcDataTester.Wpf.Interfaces;
using S7PlcDataTester.Wpf.Models;

namespace S7PlcDataTester.Wpf;

enum DataType
{
    Bool,
    Byte,
    Word,
    DWord,
    Int,
    DInt,
    Real,
    LReal,
    String
}

enum ByteValueFormat
{
    Decimal,
    Binary,
    Hexadecimal
}

public partial class MainWindow : Window
{
    private const int MaxLogLines = 1000;
    private const int MaxIpHistoryCount = 20;
    private static readonly Brush ConnectedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
    private static readonly Brush DisconnectedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C"));
    private static readonly Brush ConnectedTextBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8AF5C1"));
    private static readonly Brush DisconnectedTextBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB3C1"));
    private static readonly Brush ConnectedBadgeBackgroundBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#183629"));
    private static readonly Brush DisconnectedBadgeBackgroundBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A1B22"));
    private static readonly Brush ConnectedBadgeBorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2D8B5F"));
    private static readonly Brush DisconnectedBadgeBorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7A2A43"));
    private static readonly string IpHistoryFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PLCDataTester",
        "ip-history.json");

    private IPlcCommunicator? _plcCommunicator;
    private PlcConfig? _currentConfig;
    private readonly ObservableCollection<string> _ipHistory = new();
    private bool _isBusy;

    public MainWindow()
    {
        InitializeComponent();
        InitializeUi();
        SetupEventHandlers();
    }

    private void InitializeUi()
    {
        LoadIpHistory();
        IpAddressComboBox.ItemsSource = _ipHistory;
        IpAddressComboBox.Text = _ipHistory.FirstOrDefault() ?? "192.168.1.100";

        PlcTypeComboBox.ItemsSource = Enum.GetValues<PlcType>();
        PlcTypeComboBox.SelectedItem = PlcType.SiemensS7;

        DataTypeComboBox.ItemsSource = Enum.GetValues<DataType>();
        DataTypeComboBox.SelectedItem = DataType.Real;
        ByteFormatComboBox.ItemsSource = Enum.GetValues<ByteValueFormat>();
        ByteFormatComboBox.SelectedItem = ByteValueFormat.Decimal;
        BoolValueComboBox.ItemsSource = new[] { "TRUE", "FALSE" };
        BoolValueComboBox.SelectedIndex = 0;
        UpdateDataTypeDependentUi();

        UpdatePortBasedOnPlcType(PlcType.SiemensS7);
        UpdateAddressExampleBasedOnPlcType(PlcType.SiemensS7);
        SetConnectionUi(false);
    }

    private void SetupEventHandlers()
    {
        Closing += MainWindow_Closing;
        PlcTypeComboBox.SelectionChanged += PlcTypeComboBox_SelectionChanged;
        DataTypeComboBox.SelectionChanged += DataTypeComboBox_SelectionChanged;
        ConnectButton.Click += ConnectButton_Click;
        ReadButton.Click += ReadButton_Click;
        WriteButton.Click += WriteButton_Click;
        ClearLogButton.Click += ClearLogButton_Click;
    }

    private void DataTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateDataTypeDependentUi();
    }

    private void UpdateDataTypeDependentUi()
    {
        var selectedType = DataTypeComboBox.SelectedItem is DataType dataType ? dataType : DataType.Real;
        var isStringType = selectedType == DataType.String;
        var isByteType = selectedType == DataType.Byte;
        var isBoolType = selectedType == DataType.Bool;

        var stringVisibility = isStringType ? Visibility.Visible : Visibility.Collapsed;
        StringLengthLabel.Visibility = stringVisibility;
        StringLengthTextBox.Visibility = stringVisibility;

        ByteFormatPanel.Visibility = isByteType ? Visibility.Visible : Visibility.Collapsed;
        ValueTextBox.Visibility = isBoolType ? Visibility.Collapsed : Visibility.Visible;
        BoolValueComboBox.Visibility = isBoolType ? Visibility.Visible : Visibility.Collapsed;
    }

    private void PlcTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PlcTypeComboBox.SelectedItem is PlcType plcType)
        {
            UpdatePortBasedOnPlcType(plcType);
            UpdateAddressExampleBasedOnPlcType(plcType);
        }
    }

    private void UpdatePortBasedOnPlcType(PlcType plcType)
    {
        PortTextBox.Text = plcType switch
        {
            PlcType.SiemensS7 => "102",
            PlcType.Mitsubishi => "5000",
            PlcType.Inovance => "502",
            PlcType.KeyenceKV => "5000",
            _ => "102"
        };
    }

    private void UpdateAddressExampleBasedOnPlcType(PlcType plcType)
    {
        AddressTextBox.Text = plcType switch
        {
            PlcType.SiemensS7 => "DB1.0.0",
            PlcType.Mitsubishi => "D100",
            PlcType.Inovance => "D100",
            PlcType.KeyenceKV => "DM100",
            _ => "DB1.0.0"
        };
    }

    private async void ConnectButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        var connectStopwatch = Stopwatch.StartNew();
        ToggleBusy(true);
        try
        {
            if (_plcCommunicator?.IsConnected == true)
            {
                await DisconnectCurrentCommunicatorAsync();
                Log("Disconnected.");
                return;
            }

            if (!TryBuildConfig(out var config, out var validationError))
            {
                ResultTextBox.Text = validationError;
                Log(validationError);
                return;
            }

            await DisconnectCurrentCommunicatorAsync();
            _currentConfig = config;
            _plcCommunicator = PlcCommunicatorFactory.CreateCommunicator(config);
            _plcCommunicator.ConnectionStateChanged += PlcCommunicator_ConnectionStateChanged;

            var success = await _plcCommunicator.ConnectAsync();
            connectStopwatch.Stop();
            if (success)
            {
                RememberIpAddress(_currentConfig.IpAddress);
                SetConnectionUi(true);
                Log($"Connected to {_currentConfig.Name} ({_currentConfig.IpAddress}:{_currentConfig.Port}) in {connectStopwatch.ElapsedMilliseconds} ms.");
                ResultTextBox.Text = $"Connect succeeded ({connectStopwatch.ElapsedMilliseconds} ms).";
            }
            else
            {
                SetConnectionUi(false);
                Log($"Connect failed in {connectStopwatch.ElapsedMilliseconds} ms.");
                ResultTextBox.Text = $"Connect failed ({connectStopwatch.ElapsedMilliseconds} ms).";
            }
        }
        catch (Exception ex)
        {
            connectStopwatch.Stop();
            SetConnectionUi(false);
            ResultTextBox.Text = $"Connect failed after {connectStopwatch.ElapsedMilliseconds} ms: {ex.Message}";
            Log($"Connect failed after {connectStopwatch.ElapsedMilliseconds} ms: {ex.Message}");
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private async void ReadButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || !EnsureConnected())
        {
            return;
        }

        if (DataTypeComboBox.SelectedItem is not DataType dataType)
        {
            ResultTextBox.Text = "Please choose a data type.";
            return;
        }

        var dataTypeDescriptor = BuildDataTypeDescriptor(dataType);
        if (dataTypeDescriptor == null)
        {
            return;
        }
        dataTypeDescriptor = ResolveRuntimeDataTypeDescriptor(dataType, dataTypeDescriptor, forWrite: false);

        var address = AddressTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(address))
        {
            ResultTextBox.Text = "Address is required.";
            return;
        }

        ToggleBusy(true);
        try
        {
            var value = await _plcCommunicator!.ReadAsync(address, dataTypeDescriptor);
            if (dataType == DataType.Byte && TryConvertToByte(value, out var byteValue))
            {
                var selectedFormat = ByteFormatComboBox.SelectedItem is ByteValueFormat format ? format : ByteValueFormat.Decimal;
                var selectedText = FormatByteByPreference(byteValue, selectedFormat);
                var allFormats = BuildByteAllFormatsText(byteValue);
                ResultTextBox.Text = $"Read succeeded: {address} = {selectedText}";
                Log($"Read {address} = {selectedText} ({allFormats})");
            }
            else
            {
                ResultTextBox.Text = $"Read succeeded: {address} = {value ?? "null"}";
                Log($"Read {address} = {value ?? "null"}");
            }
        }
        catch (Exception ex)
        {
            ResultTextBox.Text = $"Read failed: {ex.Message}";
            Log($"Read failed: {ex.Message}");
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private async void WriteButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || !EnsureConnected())
        {
            return;
        }

        if (DataTypeComboBox.SelectedItem is not DataType dataType)
        {
            ResultTextBox.Text = "Please choose a data type.";
            return;
        }

        var dataTypeDescriptor = BuildDataTypeDescriptor(dataType);
        if (dataTypeDescriptor == null)
        {
            return;
        }
        dataTypeDescriptor = ResolveRuntimeDataTypeDescriptor(dataType, dataTypeDescriptor, forWrite: true);

        var address = AddressTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(address))
        {
            ResultTextBox.Text = "Address is required.";
            return;
        }

        ToggleBusy(true);
        try
        {
            var valueText = GetValueInputText(dataType);
            var value = ConvertValue(valueText, dataType);
            await _plcCommunicator!.WriteAsync(address, value, dataTypeDescriptor);
            var actualValue = await _plcCommunicator.ReadAsync(address, dataTypeDescriptor);
            var writeMatches = AreValuesEquivalent(value, actualValue, dataType);
            if (writeMatches)
            {
                ResultTextBox.Text = $"Write succeeded: {address} = {actualValue ?? "null"}";
                Log($"Write verified: {address} requested={value}, actual={actualValue ?? "null"}");
            }
            else
            {
                ResultTextBox.Text = $"Write warning: requested={value}, actual={actualValue ?? "null"}";
                Log($"Write mismatch: {address} requested={value}, actual={actualValue ?? "null"}");
            }
        }
        catch (Exception ex)
        {
            ResultTextBox.Text = $"Write failed: {ex.Message}";
            Log($"Write failed: {ex.Message}");
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private object ConvertValue(string valueStr, DataType dataType)
    {
        var culture = CultureInfo.InvariantCulture;
        return dataType switch
        {
            DataType.Bool => bool.Parse(valueStr),
            DataType.Byte => ParseByteValue(valueStr),
            DataType.Word => ushort.Parse(valueStr, NumberStyles.Integer, culture),
            DataType.DWord => uint.Parse(valueStr, NumberStyles.Integer, culture),
            DataType.Int => short.Parse(valueStr, NumberStyles.Integer, culture),
            DataType.DInt => int.Parse(valueStr, NumberStyles.Integer, culture),
            DataType.Real => float.Parse(valueStr, NumberStyles.Float | NumberStyles.AllowThousands, culture),
            DataType.LReal => double.Parse(valueStr, NumberStyles.Float | NumberStyles.AllowThousands, culture),
            DataType.String => valueStr,
            _ => throw new ArgumentOutOfRangeException(nameof(dataType), dataType, "Unknown data type.")
        };
    }

    private byte ParseByteValue(string text)
    {
        var value = text.Trim();
        var format = ByteFormatComboBox.SelectedItem is ByteValueFormat selected ? selected : ByteValueFormat.Decimal;

        return format switch
        {
            ByteValueFormat.Decimal => byte.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture),
            ByteValueFormat.Binary => Convert.ToByte(value, 2),
            ByteValueFormat.Hexadecimal => ParseHexByte(value),
            _ => byte.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture)
        };
    }

    private static byte ParseHexByte(string value)
    {
        var normalized = value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? value[2..] : value;
        return byte.Parse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
    }

    private static bool TryConvertToByte(object? value, out byte byteValue)
    {
        try
        {
            if (value is null)
            {
                byteValue = 0;
                return false;
            }

            byteValue = Convert.ToByte(value, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            byteValue = 0;
            return false;
        }
    }

    private static string BuildByteAllFormatsText(byte value)
    {
        return $"Dec={value}, Bin={Convert.ToString(value, 2).PadLeft(8, '0')}, Hex=0x{value:X2}";
    }

    private static string FormatByteByPreference(byte value, ByteValueFormat format)
    {
        return format switch
        {
            ByteValueFormat.Decimal => value.ToString(CultureInfo.InvariantCulture),
            ByteValueFormat.Binary => Convert.ToString(value, 2).PadLeft(8, '0'),
            ByteValueFormat.Hexadecimal => $"0x{value:X2}",
            _ => value.ToString(CultureInfo.InvariantCulture)
        };
    }

    private static bool AreValuesEquivalent(object expected, object? actual, DataType dataType)
    {
        if (actual is null)
        {
            return false;
        }

        return dataType switch
        {
            DataType.Real => Math.Abs(Convert.ToSingle(expected, CultureInfo.InvariantCulture) - Convert.ToSingle(actual, CultureInfo.InvariantCulture)) < 0.0001f,
            DataType.LReal => Math.Abs(Convert.ToDouble(expected, CultureInfo.InvariantCulture) - Convert.ToDouble(actual, CultureInfo.InvariantCulture)) < 0.0000001,
            DataType.String => string.Equals(Convert.ToString(expected, CultureInfo.InvariantCulture), Convert.ToString(actual, CultureInfo.InvariantCulture), StringComparison.Ordinal),
            _ => string.Equals(
                Convert.ToString(expected, CultureInfo.InvariantCulture),
                Convert.ToString(actual, CultureInfo.InvariantCulture),
                StringComparison.Ordinal)
        };
    }

    private string? BuildDataTypeDescriptor(DataType dataType)
    {
        if (dataType != DataType.String)
        {
            return dataType.ToString();
        }

        if (!int.TryParse(StringLengthTextBox.Text.Trim(), out var stringLength) || stringLength < 0 || stringLength > 1024)
        {
            ResultTextBox.Text = "String length must be in range 0-1024.";
            return null;
        }

        return $"String:{stringLength}";
    }

    private string ResolveRuntimeDataTypeDescriptor(DataType dataType, string descriptor, bool forWrite)
    {
        if (dataType != DataType.String || !descriptor.Equals("String:0", StringComparison.OrdinalIgnoreCase) || !forWrite)
        {
            return descriptor;
        }

        var runtimeLength = GetValueInputText(dataType).Length;
        if (runtimeLength <= 0)
        {
            runtimeLength = 1;
        }

        return $"String:{runtimeLength}";
    }

    private string GetValueInputText(DataType dataType)
    {
        if (dataType == DataType.Bool)
        {
            var boolText = BoolValueComboBox.SelectedItem?.ToString();
            return string.IsNullOrWhiteSpace(boolText) ? "FALSE" : boolText.Trim();
        }

        return ValueTextBox.Text.Trim();
    }

    private void PlcCommunicator_ConnectionStateChanged(object? sender, PlcConnectionEventArgs e)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(() => ApplyConnectionState(e));
            return;
        }

        ApplyConnectionState(e);
    }

    private void ApplyConnectionState(PlcConnectionEventArgs e)
    {
        SetConnectionUi(e.IsConnected);
        if (!string.IsNullOrWhiteSpace(e.Message))
        {
            Log(e.Message);
        }
    }

    private void SetConnectionUi(bool isConnected)
    {
        ConnectionStatusText.Text = isConnected ? "Connected" : "Disconnected";
        ConnectionStatusText.Foreground = isConnected ? ConnectedTextBrush : DisconnectedTextBrush;
        ConnectionStatusDot.Fill = isConnected ? ConnectedBrush : DisconnectedBrush;
        ConnectionStatusBadge.Background = isConnected ? ConnectedBadgeBackgroundBrush : DisconnectedBadgeBackgroundBrush;
        ConnectionStatusBadge.BorderBrush = isConnected ? ConnectedBadgeBorderBrush : DisconnectedBadgeBorderBrush;
        ConnectButton.Content = isConnected ? "Disconnect" : "Connect";
    }

    private bool EnsureConnected()
    {
        if (_plcCommunicator?.IsConnected == true)
        {
            return true;
        }

        ResultTextBox.Text = "Please connect PLC first.";
        Log("Please connect PLC first.");
        return false;
    }

    private bool TryBuildConfig(out PlcConfig config, out string error)
    {
        error = string.Empty;
        config = new PlcConfig();

        if (PlcTypeComboBox.SelectedItem is not PlcType plcType)
        {
            error = "Please choose PLC type.";
            return false;
        }

        var ip = IpAddressComboBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(ip))
        {
            error = "IP address is required.";
            return false;
        }

        if (!int.TryParse(PortTextBox.Text.Trim(), out var port) || port <= 0 || port > 65535)
        {
            error = "Port must be a valid number in range 1-65535.";
            return false;
        }

        config = new PlcConfig
        {
            Type = plcType,
            IpAddress = ip,
            Port = port,
            Rack = RackTextBox.Text.Trim(),
            Slot = SlotTextBox.Text.Trim(),
            Name = "PLC"
        };

        return true;
    }

    private async Task DisconnectCurrentCommunicatorAsync()
    {
        if (_plcCommunicator == null)
        {
            SetConnectionUi(false);
            return;
        }

        _plcCommunicator.ConnectionStateChanged -= PlcCommunicator_ConnectionStateChanged;
        if (_plcCommunicator.IsConnected)
        {
            await _plcCommunicator.DisconnectAsync();
        }

        _plcCommunicator = null;
        SetConnectionUi(false);
    }

    private void Log(string message)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(() => Log(message));
            return;
        }

        if (LogTextBox.LineCount > MaxLogLines)
        {
            LogTextBox.Clear();
            LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] Log cleared automatically (>{MaxLogLines} lines).\n");
        }

        LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
        LogTextBox.ScrollToEnd();
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        LogTextBox.Clear();
        Log("Log cleared.");
    }

    private async void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            await DisconnectCurrentCommunicatorAsync();
        }
        catch
        {
            // ignore close-time cleanup failures
        }
    }

    private void ToggleBusy(bool isBusy)
    {
        _isBusy = isBusy;
        ConnectButton.IsEnabled = !isBusy;
        ReadButton.IsEnabled = !isBusy;
        WriteButton.IsEnabled = !isBusy;
    }

    private void LoadIpHistory()
    {
        try
        {
            if (!File.Exists(IpHistoryFilePath))
            {
                return;
            }

            var json = File.ReadAllText(IpHistoryFilePath);
            var values = JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            foreach (var ip in values.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(MaxIpHistoryCount))
            {
                _ipHistory.Add(ip);
            }
        }
        catch
        {
            // ignore malformed history file
        }
    }

    private void RememberIpAddress(string ip)
    {
        var normalized = ip.Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        var existing = _ipHistory.FirstOrDefault(x => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            _ipHistory.Remove(existing);
        }

        _ipHistory.Insert(0, normalized);
        while (_ipHistory.Count > MaxIpHistoryCount)
        {
            _ipHistory.RemoveAt(_ipHistory.Count - 1);
        }

        SaveIpHistory();
    }

    private void SaveIpHistory()
    {
        try
        {
            var directory = Path.GetDirectoryName(IpHistoryFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(_ipHistory.ToList());
            File.WriteAllText(IpHistoryFilePath, json);
        }
        catch
        {
            // ignore history persistence failures
        }
    }
}

