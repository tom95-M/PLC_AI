using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace S7PlcSimulator;

public enum PlcValueType
{
    Array,
    Bool,
    Byte,
    Usint,
    Sint,
    Word,
    Uint,
    Dword,
    Udint,
    Char,
    Int,
    Dint,
    Real,
    LReal,
    S5Time,
    Time,
    Date,
    TimeOfDay,
    DateAndTime,
    String,
    WString
}

public sealed class VariableRow : INotifyPropertyChanged
{
    private const int DefaultStringLength = 256;
    private string _name = "Var";
    private int _dbNumber = 1;
    private int _offset;
    private int _bit;
    private PlcValueType _type = PlcValueType.Bool;
    private int _stringLength = 256;
    private bool _isStringLengthExplicit;
    private PlcValueType _arrayElementType = PlcValueType.Bool;
    private int _arrayLowerBound;
    private int _arrayUpperBound = 1;
    private string _startValue = "FALSE";
    private string _value = "false";
    private bool _isArrayExpanded = true;
    private bool _isSyncingArrayItems;

    public VariableRow()
    {
        ArrayItems.CollectionChanged += (_, _) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayItems)));
        RebuildArrayItems();
    }

    public string Name
    {
        get => _name;
        set
        {
            if (EqualityComparer<string>.Default.Equals(_name, value))
            {
                return;
            }

            _name = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
            RefreshArrayItemMetadata();
        }
    }

    public int DbNumber
    {
        get => _dbNumber;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_dbNumber, value))
            {
                return;
            }

            _dbNumber = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DbNumber)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RefreshArrayItemMetadata();
        }
    }

    public int Offset
    {
        get => _offset;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_offset, value))
            {
                return;
            }

            _offset = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Offset)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OffsetDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RefreshArrayItemMetadata();
        }
    }

    public int Bit
    {
        get => _bit;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_bit, value))
            {
                return;
            }

            _bit = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bit)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OffsetDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
        }
    }

    public PlcValueType Type
    {
        get => _type;
        set
        {
            if (EqualityComparer<PlcValueType>.Default.Equals(_type, value))
            {
                return;
            }

            _type = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Type)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeLabel)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayFirstValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayFirstMonitorValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MonitorValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OffsetDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RebuildArrayItems();
        }
    }

    public string StartValue
    {
        get => _startValue;
        set
        {
            if (EqualityComparer<string>.Default.Equals(_startValue, value))
            {
                return;
            }

            _startValue = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartValue)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayFirstValueDisplay)));
            if (!_isSyncingArrayItems)
            {
                SyncArrayItemsFromStartValue();
            }
        }
    }

    public int StringLength
    {
        get => _stringLength;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_stringLength, value))
            {
                return;
            }

            _stringLength = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StringLength)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OffsetDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
        }
    }

    public string Value
    {
        get => _value;
        set
        {
            if (EqualityComparer<string>.Default.Equals(_value, value))
            {
                return;
            }

            _value = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MonitorValueDisplay)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayFirstMonitorValueDisplay)));
            if (!_isSyncingArrayItems)
            {
                SyncArrayItemsFromMonitorValue();
            }
        }
    }

    public PlcValueType ArrayElementType
    {
        get => _arrayElementType;
        set
        {
            if (EqualityComparer<PlcValueType>.Default.Equals(_arrayElementType, value))
            {
                return;
            }

            _arrayElementType = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayElementType)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RebuildArrayItems();
        }
    }

    public int ArrayLowerBound
    {
        get => _arrayLowerBound;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_arrayLowerBound, value))
            {
                return;
            }

            _arrayLowerBound = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayLowerBound)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayLength)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RebuildArrayItems();
        }
    }

    public int ArrayUpperBound
    {
        get => _arrayUpperBound;
        set
        {
            if (EqualityComparer<int>.Default.Equals(_arrayUpperBound, value))
            {
                return;
            }

            _arrayUpperBound = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayUpperBound)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayLength)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
            RebuildArrayItems();
        }
    }

    public int ArrayLength => Math.Max(1, ArrayUpperBound - ArrayLowerBound + 1);

    public bool IsArrayExpanded
    {
        get => _isArrayExpanded;
        set => SetField(ref _isArrayExpanded, value);
    }

    public ObservableCollection<ArrayElementItem> ArrayItems { get; } = [];

    public string OffsetDisplay => Type == PlcValueType.Bool
        ? $"{Offset}.{Bit}"
        : $"{Offset}.0";

    public string ArrayFirstValueDisplay
    {
        get => Type == PlcValueType.Array ? ExtractFirstArrayItem(StartValue) : StartValue;
    }

    public string ArrayFirstMonitorValueDisplay
    {
        get => Type == PlcValueType.Array ? ExtractFirstArrayItem(Value) : Value;
    }

    public string StartValueDisplay => Type == PlcValueType.Array ? string.Empty : StartValue;

    public string MonitorValueDisplay => Type == PlcValueType.Array ? string.Empty : Value;

    public string AddressDisplay => Type switch
    {
        PlcValueType.Array => BuildArrayStartAddress(DbNumber, Offset, ArrayElementType),
        PlcValueType.Bool => $"DB{DbNumber}.DBX{Offset}.{Bit}",
        PlcValueType.Byte => $"DB{DbNumber}.DBB{Offset}",
        PlcValueType.Usint => $"DB{DbNumber}.DBB{Offset}",
        PlcValueType.Sint => $"DB{DbNumber}.DBB{Offset}",
        PlcValueType.Char => $"DB{DbNumber}.DBB{Offset}",
        PlcValueType.Word => $"DB{DbNumber}.DBW{Offset}",
        PlcValueType.Uint => $"DB{DbNumber}.DBW{Offset}",
        PlcValueType.Int => $"DB{DbNumber}.DBW{Offset}",
        PlcValueType.Date => $"DB{DbNumber}.DBW{Offset}",
        PlcValueType.S5Time => $"DB{DbNumber}.DBW{Offset}",
        PlcValueType.Dword => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.Udint => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.Dint => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.Real => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.LReal => $"DB{DbNumber}.DBD{Offset}[8]",
        PlcValueType.Time => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.TimeOfDay => $"DB{DbNumber}.DBD{Offset}",
        PlcValueType.DateAndTime => $"DB{DbNumber}.DBB{Offset}[8]",
        PlcValueType.String => $"DB{DbNumber}.STRING{Offset}[{StringLength}]",
        PlcValueType.WString => $"DB{DbNumber}.WSTRING{Offset}[{StringLength}]",
        _ => $"DB{DbNumber}@{Offset}"
    };

    public string TypeLabel => Type switch
    {
        PlcValueType.Array => "Array",
        PlcValueType.Bool => "Bool",
        PlcValueType.Byte => "Byte",
        PlcValueType.Usint => "Usint",
        PlcValueType.Sint => "Sint",
        PlcValueType.Word => "Word",
        PlcValueType.Uint => "Uint",
        PlcValueType.Dword => "Dword",
        PlcValueType.Udint => "Udint",
        PlcValueType.Char => "Char",
        PlcValueType.Int => "Int",
        PlcValueType.Dint => "Dint",
        PlcValueType.Real => "Real",
        PlcValueType.LReal => "LReal",
        PlcValueType.S5Time => "S5Time",
        PlcValueType.Time => "Time",
        PlcValueType.Date => "Date",
        PlcValueType.TimeOfDay => "TimeOfDay",
        PlcValueType.DateAndTime => "DateAndTime",
        PlcValueType.String => "String",
        PlcValueType.WString => "WString",
        _ => Type.ToString()
    };

    public string TypeEditorText
    {
        get => Type switch
        {
            PlcValueType.Array when ArrayElementType == PlcValueType.String => $"Array[{ArrayLowerBound}..{ArrayUpperBound}] of String[{StringLength}]",
            PlcValueType.Array => $"Array[{ArrayLowerBound}..{ArrayUpperBound}] of {ArrayElementType}",
            PlcValueType.String when _isStringLengthExplicit => $"{TypeLabel}[{StringLength}]",
            PlcValueType.String => TypeLabel,
            PlcValueType.WString => $"{TypeLabel}[{StringLength}]",
            _ => TypeLabel
        };
        set
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            var text = value.Trim();
            if (TryParseArrayTypeEditorText(text, out var lowerBound, out var upperBound, out var elementType, out var arrayStringLength))
            {
                ArrayLowerBound = lowerBound;
                ArrayUpperBound = upperBound;
                ArrayElementType = elementType;
                if (elementType == PlcValueType.String && arrayStringLength.HasValue)
                {
                    StringLength = arrayStringLength.Value;
                }
                _isStringLengthExplicit = false;
                Type = PlcValueType.Array;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
                return;
            }

            var match = Regex.Match(text, @"^([A-Za-z][A-Za-z0-9]*)(?:\[(\d+)\])?$");
            if (!match.Success)
            {
                return;
            }

            var rawType = match.Groups[1].Value;
            var rawLength = match.Groups[2].Success ? match.Groups[2].Value : null;
            if (!TryParseType(rawType, out var parsedType))
            {
                return;
            }

            if (parsedType is PlcValueType.String or PlcValueType.WString)
            {
                if (!string.IsNullOrEmpty(rawLength) && int.TryParse(rawLength, out var parsedLength) && parsedLength > 0)
                {
                    StringLength = parsedLength;
                    if (parsedType == PlcValueType.String)
                    {
                        _isStringLengthExplicit = true;
                    }
                }
                else if (parsedType == PlcValueType.String)
                {
                    StringLength = DefaultStringLength;
                    _isStringLengthExplicit = false;
                }
            }
            else
            {
                _isStringLengthExplicit = false;
                if (parsedType == PlcValueType.Array)
                {
                    ArrayLowerBound = 0;
                    ArrayUpperBound = 1;
                    ArrayElementType = PlcValueType.Bool;
                }
            }

            Type = parsedType;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeEditorText)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static bool TryParseType(string rawType, out PlcValueType type)
    {
        var normalized = rawType.Trim().ToLowerInvariant();
        return normalized switch
        {
            "array" => SetType(PlcValueType.Array, out type),
            "bool" => SetType(PlcValueType.Bool, out type),
            "byte" => SetType(PlcValueType.Byte, out type),
            "usint" => SetType(PlcValueType.Usint, out type),
            "sint" => SetType(PlcValueType.Sint, out type),
            "word" => SetType(PlcValueType.Word, out type),
            "uint" => SetType(PlcValueType.Uint, out type),
            "dword" => SetType(PlcValueType.Dword, out type),
            "udint" => SetType(PlcValueType.Udint, out type),
            "char" => SetType(PlcValueType.Char, out type),
            "int" => SetType(PlcValueType.Int, out type),
            "dint" => SetType(PlcValueType.Dint, out type),
            "real" => SetType(PlcValueType.Real, out type),
            "lreal" => SetType(PlcValueType.LReal, out type),
            "s5time" => SetType(PlcValueType.S5Time, out type),
            "time" => SetType(PlcValueType.Time, out type),
            "date" => SetType(PlcValueType.Date, out type),
            "timeofday" => SetType(PlcValueType.TimeOfDay, out type),
            "dateandtime" => SetType(PlcValueType.DateAndTime, out type),
            "string" => SetType(PlcValueType.String, out type),
            "wstring" => SetType(PlcValueType.WString, out type),
            _ => SetType(default, out type, false)
        };
    }

    private static bool TryParseArrayTypeEditorText(string text, out int lowerBound, out int upperBound, out PlcValueType elementType, out int? stringLength)
    {
        lowerBound = 0;
        upperBound = 0;
        elementType = default;
        stringLength = null;

        var match = Regex.Match(text, @"^array\[\s*(-?\d+)\s*\.\.\s*(-?\d+)\s*\]\s*of\s*([A-Za-z][A-Za-z0-9]*)(?:\[(\d+)\])?$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return false;
        }

        if (!int.TryParse(match.Groups[1].Value, out lowerBound) ||
            !int.TryParse(match.Groups[2].Value, out upperBound) ||
            upperBound < lowerBound)
        {
            return false;
        }

        if (!TryParseType(match.Groups[3].Value, out elementType))
        {
            return false;
        }

        if (elementType == PlcValueType.String)
        {
            if (!match.Groups[4].Success || !int.TryParse(match.Groups[4].Value, out var parsedLength) || parsedLength is < 1 or > 254)
            {
                return false;
            }

            stringLength = parsedLength;
        }

        return IsArrayElementTypeSupported(elementType);
    }

    private static bool IsArrayElementTypeSupported(PlcValueType type)
    {
        return type switch
        {
            PlcValueType.Array => false,
            PlcValueType.WString => false,
            _ => true
        };
    }

    private static bool SetType(PlcValueType value, out PlcValueType type, bool ok = true)
    {
        type = value;
        return ok;
    }

    private static string BuildArrayStartAddress(int dbNumber, int offset, PlcValueType elementType)
    {
        return elementType switch
        {
            PlcValueType.Bool => $"DB{dbNumber}.DBX{offset}.0",
            PlcValueType.Byte => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Usint => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Sint => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Char => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Word => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Uint => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Int => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Date => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.S5Time => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Dword => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Udint => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Dint => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Real => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Time => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.TimeOfDay => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.LReal => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.DateAndTime => $"DB{dbNumber}.DBB{offset}",
            _ => $"DB{dbNumber}@{offset}"
        };
    }

    private void RebuildArrayItems()
    {
        if (Type != PlcValueType.Array)
        {
            if (ArrayItems.Count == 0)
            {
                return;
            }

            foreach (var item in ArrayItems)
            {
                item.PropertyChanged -= ArrayItem_OnPropertyChanged;
            }

            ArrayItems.Clear();
            return;
        }

        var startValues = ParseArrayValues(StartValue, ArrayLength, GetDefaultArrayElementValue(ArrayElementType));
        var monitorValues = ParseArrayValues(Value, ArrayLength, GetDefaultArrayElementValue(ArrayElementType));
        _isSyncingArrayItems = true;
        try
        {
            foreach (var item in ArrayItems)
            {
                item.PropertyChanged -= ArrayItem_OnPropertyChanged;
            }

            ArrayItems.Clear();
            for (var i = 0; i < ArrayLength; i++)
            {
                var elementOffset = Offset + i * GetScalarElementSize(ArrayElementType);
                var item = new ArrayElementItem(
                    Name,
                    ArrayLowerBound + i,
                    ArrayElementType,
                    DbNumber,
                    elementOffset,
                    startValues[i],
                    monitorValues[i]);
                item.PropertyChanged += ArrayItem_OnPropertyChanged;
                ArrayItems.Add(item);
            }
        }
        finally
        {
            _isSyncingArrayItems = false;
        }

        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ArrayItems)));
    }

    private void SyncArrayItemsFromStartValue()
    {
        if (Type != PlcValueType.Array)
        {
            return;
        }

        if (ArrayItems.Count != ArrayLength)
        {
            RebuildArrayItems();
            return;
        }

        var values = ParseArrayValues(StartValue, ArrayLength, GetDefaultArrayElementValue(ArrayElementType));
        _isSyncingArrayItems = true;
        try
        {
            for (var i = 0; i < ArrayItems.Count; i++)
            {
                ArrayItems[i].StartValue = values[i];
            }
        }
        finally
        {
            _isSyncingArrayItems = false;
        }
    }

    private void SyncArrayItemsFromMonitorValue()
    {
        if (Type != PlcValueType.Array)
        {
            return;
        }

        if (ArrayItems.Count != ArrayLength)
        {
            RebuildArrayItems();
            return;
        }

        var values = ParseArrayValues(Value, ArrayLength, GetDefaultArrayElementValue(ArrayElementType));
        _isSyncingArrayItems = true;
        try
        {
            for (var i = 0; i < ArrayItems.Count; i++)
            {
                ArrayItems[i].Value = values[i];
            }
        }
        finally
        {
            _isSyncingArrayItems = false;
        }
    }

    private void RefreshArrayItemMetadata()
    {
        if (Type != PlcValueType.Array || ArrayItems.Count == 0)
        {
            return;
        }

        var elementSize = GetScalarElementSize(ArrayElementType);
        for (var i = 0; i < ArrayItems.Count; i++)
        {
            ArrayItems[i].UpdateMeta(Name, ArrayLowerBound + i, ArrayElementType, DbNumber, Offset + i * elementSize);
        }
    }

    private void ArrayItem_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_isSyncingArrayItems || e.PropertyName != nameof(ArrayElementItem.StartValue))
        {
            return;
        }

        _isSyncingArrayItems = true;
        try
        {
            StartValue = $"[{string.Join(", ", ArrayItems.Select(x => x.StartValue))}]";
        }
        finally
        {
            _isSyncingArrayItems = false;
        }
    }

    private static string[] ParseArrayValues(string raw, int expectedCount, string fallback)
    {
        var value = raw?.Trim() ?? string.Empty;
        if (value.StartsWith('[') && value.EndsWith(']') && value.Length >= 2)
        {
            value = value[1..^1].Trim();
        }

        var parts = string.IsNullOrWhiteSpace(value)
            ? []
            : value.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        var result = new string[expectedCount];
        for (var i = 0; i < expectedCount; i++)
        {
            result[i] = i < parts.Length ? parts[i] : fallback;
        }
        return result;
    }

    private static string GetDefaultArrayElementValue(PlcValueType type)
    {
        return type switch
        {
            PlcValueType.Bool => "FALSE",
            PlcValueType.Date => "1990-01-01",
            PlcValueType.TimeOfDay => "00:00:00.000",
            PlcValueType.DateAndTime => "1990-01-01 00:00:00.000",
            _ => "0"
        };
    }

    private static int GetScalarElementSize(PlcValueType type)
    {
        return type switch
        {
            PlcValueType.Bool => 1,
            PlcValueType.Byte => 1,
            PlcValueType.Usint => 1,
            PlcValueType.Sint => 1,
            PlcValueType.Char => 1,
            PlcValueType.Word => 2,
            PlcValueType.Uint => 2,
            PlcValueType.Int => 2,
            PlcValueType.Date => 2,
            PlcValueType.S5Time => 2,
            PlcValueType.Dword => 4,
            PlcValueType.Udint => 4,
            PlcValueType.Dint => 4,
            PlcValueType.Real => 4,
            PlcValueType.Time => 4,
            PlcValueType.TimeOfDay => 4,
            PlcValueType.LReal => 8,
            PlcValueType.DateAndTime => 8,
            _ => 1
        };
    }

    private static string ExtractFirstArrayItem(string? raw)
    {
        var value = raw?.Trim() ?? string.Empty;
        if (value.StartsWith('[') && value.EndsWith(']') && value.Length >= 2)
        {
            value = value[1..^1];
        }

        var first = value.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return first ?? string.Empty;
    }
}

public sealed class ArrayElementItem : INotifyPropertyChanged
{
    private string _parentName;
    private int _index;
    private PlcValueType _elementType;
    private int _dbNumber;
    private int _offset;
    private string _startValue;
    private string _value;

    public ArrayElementItem(string parentName, int index, PlcValueType elementType, int dbNumber, int offset, string startValue, string value)
    {
        _parentName = parentName;
        _index = index;
        _elementType = elementType;
        _dbNumber = dbNumber;
        _offset = offset;
        _startValue = startValue;
        _value = value;
    }

    public string NameDisplay => $"{_parentName}[{_index}]";
    public string TypeDisplay => _elementType.ToString();
    public PlcValueType Type => _elementType;
    public bool IsBool => _elementType == PlcValueType.Bool;
    public string OffsetDisplay => _elementType == PlcValueType.Bool ? $"{_offset}.0" : $"{_offset}.0";
    public string AddressDisplay => BuildAddressDisplay(_dbNumber, _offset, _elementType);

    public string StartValue
    {
        get => _startValue;
        set
        {
            if (EqualityComparer<string>.Default.Equals(_startValue, value))
            {
                return;
            }

            _startValue = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartValue)));
        }
    }

    public string Value
    {
        get => _value;
        set
        {
            if (EqualityComparer<string>.Default.Equals(_value, value))
            {
                return;
            }

            _value = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void UpdateMeta(string parentName, int index, PlcValueType elementType, int dbNumber, int offset)
    {
        _parentName = parentName;
        _index = index;
        _elementType = elementType;
        _dbNumber = dbNumber;
        _offset = offset;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NameDisplay)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeDisplay)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Type)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBool)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OffsetDisplay)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AddressDisplay)));
    }

    private static string BuildAddressDisplay(int dbNumber, int offset, PlcValueType type)
    {
        return type switch
        {
            PlcValueType.Bool => $"DB{dbNumber}.DBX{offset}.0",
            PlcValueType.Byte => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Usint => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Sint => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Char => $"DB{dbNumber}.DBB{offset}",
            PlcValueType.Word => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Uint => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Int => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Date => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.S5Time => $"DB{dbNumber}.DBW{offset}",
            PlcValueType.Dword => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Udint => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Dint => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Real => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.Time => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.TimeOfDay => $"DB{dbNumber}.DBD{offset}",
            PlcValueType.LReal => $"DB{dbNumber}.DBD{offset}[8]",
            PlcValueType.DateAndTime => $"DB{dbNumber}.DBB{offset}[8]",
            _ => $"DB{dbNumber}@{offset}"
        };
    }
}
