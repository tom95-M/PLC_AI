﻿using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Buffers.Binary;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace S7PlcSimulator;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private const string AppStateFilePath = "simulator_state.json";
    private const int DefaultDbSize = 1000;
    private const int DefaultStringLength = 256;
    private const int DefaultPlcPort = 102;
    private static readonly DateOnly S7DateBase = new(1990, 1, 1);
    private PlcMemoryStore _memoryStore;
    private readonly Dictionary<ProjectTreeNode, S7ServerHost> _plcServerHosts = [];
    private readonly Dictionary<ProjectTreeNode, PlcEndpointSettings> _plcEndpointSettings = [];
    private bool _isRefreshingValues;
    private bool _isMonitorEnabled;
    private bool _isMonitorValueCellContext;
    private int? _activeDbNumber;
    private ProjectTreeNode? _selectedNode;
    private readonly DispatcherTimer _valueRefreshTimer = new();
    private readonly Dictionary<VariableRow, string> _typeEditorTextBeforeEdit = [];
    private const int MaxArrayElementCount = 10_000;

    public ObservableCollection<VariableRow> Variables { get; } = new();
    public ObservableCollection<ProjectTreeNode> ProjectNodes { get; } = new();
    public IReadOnlyList<PlcTypeOption> TypeOptions { get; } =
    [
        new PlcTypeOption(PlcValueType.Array, "Array[0..1] of Bool"),
        new PlcTypeOption(PlcValueType.Bool, "Bool"),
        new PlcTypeOption(PlcValueType.Byte, "Byte"),
        new PlcTypeOption(PlcValueType.Usint, "USInt"),
        new PlcTypeOption(PlcValueType.Sint, "SInt"),
        new PlcTypeOption(PlcValueType.Word, "Word"),
        new PlcTypeOption(PlcValueType.Uint, "Uint"),
        new PlcTypeOption(PlcValueType.Dword, "Dword"),
        new PlcTypeOption(PlcValueType.Udint, "UDint"),
        new PlcTypeOption(PlcValueType.Char, "Char"),
        new PlcTypeOption(PlcValueType.Int, "Int"),
        new PlcTypeOption(PlcValueType.Dint, "Dint"),
        new PlcTypeOption(PlcValueType.Real, "Real"),
        new PlcTypeOption(PlcValueType.LReal, "LReal"),
        new PlcTypeOption(PlcValueType.S5Time, "S5Time"),
        new PlcTypeOption(PlcValueType.Time, "Time"),
        new PlcTypeOption(PlcValueType.Date, "Date"),
        new PlcTypeOption(PlcValueType.TimeOfDay, "TimeOfDay"),
        new PlcTypeOption(PlcValueType.DateAndTime, "DateAndTime"),
        new PlcTypeOption(PlcValueType.String, "String"),
        new PlcTypeOption(PlcValueType.WString, "WString")
    ];
    public ICollectionView FilteredVariables { get; }
    public bool IsMonitorEnabled
    {
        get => _isMonitorEnabled;
        private set
        {
            if (_isMonitorEnabled == value)
            {
                return;
            }

            _isMonitorEnabled = value;
            OnPropertyChanged(nameof(IsMonitorEnabled));
        }
    }
    public string CurrentDataBlockName => _selectedNode?.NodeType == ProjectNodeType.DataBlock
        ? _selectedNode.Name
        : "变量面板";
    public bool IsPlcOverviewSelected => _selectedNode?.NodeType == ProjectNodeType.PlcRoot;
    public bool IsProgramAreaSelected => _selectedNode?.NodeType is ProjectNodeType.ProgramFolder or ProjectNodeType.ProgramGroup;
    public bool IsDataBlockSelected => _selectedNode?.NodeType == ProjectNodeType.DataBlock;
    public bool IsVariablePanelVisible => !IsPlcOverviewSelected && !IsProgramAreaSelected;

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainWindow()
    {
        FilteredVariables = CollectionViewSource.GetDefaultView(Variables);
        FilteredVariables.Filter = item => item is VariableRow row && _activeDbNumber.HasValue && row.DbNumber == _activeDbNumber.Value;
        EnsureVariablesSortOrder();
        if (FilteredVariables is ICollectionViewLiveShaping liveView && liveView.CanChangeLiveSorting)
        {
            liveView.LiveSortingProperties.Add(nameof(VariableRow.Offset));
            liveView.LiveSortingProperties.Add(nameof(VariableRow.Bit));
            liveView.IsLiveSorting = true;
        }

        InitializeComponent();
        DataContext = this;

        var appState = LoadAppState(AppStateFilePath);
        var dbConfig = appState is not null
            ? BuildDbConfigFromState(appState)
            : (LoadDbConfig("dbconfig.json") ?? new Dictionary<int, int> { [1] = 1000 });
        _memoryStore = new PlcMemoryStore(dbConfig);

        var restored = appState is not null && TryRestoreAppState(appState);
        if (!restored)
        {
            InitializeProjectTree();

            Variables.Add(new VariableRow { Name = "MotorRun", DbNumber = 1, Offset = 0, Bit = 0, Type = PlcValueType.Bool, StringLength = DefaultStringLength, StartValue = "false", Value = "false" });
            Variables.Add(new VariableRow { Name = "Counter", DbNumber = 1, Offset = 2, Type = PlcValueType.Int, StringLength = DefaultStringLength, StartValue = "100", Value = "100" });
            Variables.Add(new VariableRow { Name = "Pressure", DbNumber = 1, Offset = 4, Type = PlcValueType.Real, StringLength = DefaultStringLength, StartValue = "3.14", Value = "3.14" });
        }

        foreach (var row in Variables)
        {
            row.PropertyChanged += VariableRow_OnPropertyChanged;
        }
        Variables.CollectionChanged += Variables_OnCollectionChanged;
        foreach (var db in Variables.Select(x => x.DbNumber).Distinct().ToList())
        {
            RecalculateOffsetsForDb(db);
        }
        _selectedNode = ProjectNodes.FirstOrDefault(x => x.NodeType == ProjectNodeType.PlcRoot)
                        ?? _selectedNode
                        ?? FindFirstDataBlockNode()
                        ?? ProjectNodes.FirstOrDefault();
        RefreshVariableFilter();

        _valueRefreshTimer.Interval = TimeSpan.FromMilliseconds(200);
        _valueRefreshTimer.Tick += ValueRefreshTimer_OnTick;
        SetMonitorEnabled(false, writeLog: false);

        UpdateConnectionStatus(false);
        UpdateSimulatorRunState();
        AddLog("界面已初始化，可手动编辑程序组、数据块和变量。");
    }

    protected override void OnClosed(EventArgs e)
    {
        SaveAppState(AppStateFilePath);
        foreach (var host in _plcServerHosts.Values)
        {
            host.Stop();
        }
        _plcServerHosts.Clear();
        base.OnClosed(e);
    }

    private void InitializeProjectTree()
    {
        AddPlcSimulator("S7-1200/S7-1500模拟器_1");
    }

    private void ProjectTreeView_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        _selectedNode = e.NewValue as ProjectTreeNode;
        RefreshVariableFilter();
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is not null && _plcEndpointSettings.TryGetValue(plcRoot, out var settings))
        {
            AddressTextBox.Text = settings.Address;
            PortTextBox.Text = settings.Port.ToString(CultureInfo.InvariantCulture);
        }
    }

    private void ProjectTreeView_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var item = FindVisualParent<TreeViewItem>(e.OriginalSource as DependencyObject);
        if (item is null)
        {
            return;
        }

        item.IsSelected = true;
        item.Focus();
    }

    private void ProjectTreeView_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_selectedNode is null)
        {
            ProjectTreeView.ContextMenu = null;
            return;
        }

        var menu = new ContextMenu();
        ApplyContextMenuStyle(menu);
        menu.Items.Add(CreateMenuItem("重命名", RenameNodeButton_OnClick));

        switch (_selectedNode.NodeType)
        {
            case ProjectNodeType.PlcRoot:
                menu.Items.Add(CreateMenuItem("新增PLC模拟器", AddPlcSimulatorButton_OnClick));
                menu.Items.Add(CreateMenuItem("删除PLC模拟器", RemovePlcSimulatorButton_OnClick));
                menu.Items.Add(CreateMenuItem("PLC设置", PlcSettingsButton_OnClick));
                break;

            case ProjectNodeType.ProgramFolder:
            case ProjectNodeType.ProgramGroup:
                menu.Items.Add(CreateMenuItem("新增数据块", AddDbBlockButton_OnClick));
                break;

            case ProjectNodeType.DataBlockFolder:
            case ProjectNodeType.DataBlock:
                menu.Items.Add(CreateMenuItem("新增数据块", AddDbBlockButton_OnClick));
                if (_selectedNode.NodeType == ProjectNodeType.DataBlock)
                {
                    menu.Items.Add(CreateMenuItem("删除数据块", RemoveDbBlockButton_OnClick));
                }
                break;
        }

        if (_selectedNode.NodeType == ProjectNodeType.DataBlock)
        {
            menu.Items.Add(CreateContextMenuSeparator());
            menu.Items.Add(CreateMenuItem("属性", DbBlockPropertiesButton_OnClick));
        }

        ProjectTreeView.ContextMenu = menu;
    }

    private void AddPlcSimulatorButton_OnClick(object sender, RoutedEventArgs e)
    {
        var nextIndex = FindNextPlcIndex();
        AddPlcSimulator($"S7-1200/S7-1500模拟器_{nextIndex}");
    }

    private void RemovePlcSimulatorButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_selectedNode is null || _selectedNode.NodeType != ProjectNodeType.PlcRoot)
        {
            AddLog("请先选中要删除的 PLC 模拟器。");
            return;
        }

        if (ProjectNodes.Count <= 1)
        {
            AddLog("至少保留一个 PLC 模拟器，不能全部删除。");
            return;
        }

        var removedName = _selectedNode.Name;
        var removeIndex = ProjectNodes.IndexOf(_selectedNode);
        StopPlcServer(_selectedNode);
        _plcEndpointSettings.Remove(_selectedNode);
        ProjectNodes.Remove(_selectedNode);
        SortPlcRootsByIndex();

        if (ProjectNodes.Count > 0)
        {
            var nextIndex = Math.Min(removeIndex, ProjectNodes.Count - 1);
            _selectedNode = ProjectNodes[nextIndex];
        }
        else
        {
            _selectedNode = null;
        }

        AddLog($"已删除 PLC 模拟器: {removedName}");
        RefreshVariableFilter();
        UpdateSimulatorRunState();
    }

    private void PlcSettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_selectedNode is null || _selectedNode.NodeType != ProjectNodeType.PlcRoot)
        {
            return;
        }

        var current = GetOrCreatePlcSettings(_selectedNode);
        var updated = PromptForPlcSettings(_selectedNode, current);
        if (updated is null)
        {
            return;
        }

        _plcEndpointSettings[_selectedNode] = updated;
        ApplyPlcSettings(_selectedNode, updated);
    }

    private void ThemeSettingsMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        var combo = new ComboBox
        {
            MinWidth = 220,
            Margin = new Thickness(0, 8, 0, 12),
            ItemsSource = new[] { "浅色", "深色" },
            SelectedIndex = 0
        };

        var windowBrush = Resources["WindowBrush"] as SolidColorBrush;
        if (windowBrush is not null && windowBrush.Color == (Color)ColorConverter.ConvertFromString("#1C2430"))
        {
            combo.SelectedItem = "深色";
        }

        var okButton = new Button { Content = "确定", Width = 84, IsDefault = true };
        var cancelButton = new Button { Content = "取消", Width = 84, IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
        var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock { Text = "请选择主题：" });
        root.Children.Add(combo);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = "主题设置",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        okButton.Click += (_, _) =>
        {
            ApplyTheme((combo.SelectedItem as string) == "深色" ? "dark" : "light");
            dialog.DialogResult = true;
            dialog.Close();
        };

        dialog.ShowDialog();
    }

    private void IpTestMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        var selectedPlc = GetCurrentPlcRoot();
        var defaultIp = selectedPlc is not null && _plcEndpointSettings.TryGetValue(selectedPlc, out var selectedSettings)
            ? selectedSettings.Address
            : string.Empty;

        var ipInput = new TextBox { Text = defaultIp, MinWidth = 220, Margin = new Thickness(0, 8, 0, 10) };
        var portInput = new TextBox { Text = DefaultPlcPort.ToString(CultureInfo.InvariantCulture), MinWidth = 220, Margin = new Thickness(0, 0, 0, 12) };
        var testButton = new Button { Content = "测试", Width = 84, IsDefault = true };
        var closeButton = new Button { Content = "关闭", Width = 84, IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
        var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttonPanel.Children.Add(testButton);
        buttonPanel.Children.Add(closeButton);

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock { Text = "IP 地址：" });
        root.Children.Add(ipInput);
        root.Children.Add(new TextBlock { Text = "端口：" });
        root.Children.Add(portInput);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = "IP测试",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        testButton.Click += async (_, _) =>
        {
            var ipText = ipInput.Text.Trim();
            if (!IPAddress.TryParse(ipText, out var ip) || ip.AddressFamily != AddressFamily.InterNetwork)
            {
                MessageBox.Show(dialog, "请输入合法的客户端 IPv4 地址。", "IP测试", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(portInput.Text.Trim(), out var port) || port is < 1 or > 65535)
            {
                MessageBox.Show(dialog, "端口范围必须在 1-65535。", "IP测试", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            testButton.IsEnabled = false;
            try
            {
                string pingResult;
                var pingOk = false;
                try
                {
                    using var ping = new Ping();
                    var reply = await ping.SendPingAsync(ip, 1200);
                    pingOk = reply.Status == IPStatus.Success;
                    pingResult = pingOk
                        ? $"Ping 成功，延迟 {reply.RoundtripTime} ms"
                        : $"Ping 失败：{reply.Status}";
                }
                catch (Exception ex)
                {
                    pingResult = $"Ping 异常：{ex.Message}";
                }

                string tcpResult;
                var tcpOk = false;
                try
                {
                    using var client = new TcpClient();
                    var connectTask = client.ConnectAsync(ip, port);
                    var timeoutTask = Task.Delay(1500);
                    var completed = await Task.WhenAny(connectTask, timeoutTask);
                    if (completed == connectTask)
                    {
                        await connectTask;
                        tcpOk = true;
                        tcpResult = "TCP 连接成功";
                    }
                    else
                    {
                        tcpResult = "TCP 连接超时";
                    }
                }
                catch (Exception ex)
                {
                    tcpResult = $"TCP 连接失败：{ex.Message}";
                }

                var summary = tcpOk || pingOk ? "通断测试：可达" : "通断测试：不可达";
                var icon = tcpOk || pingOk ? MessageBoxImage.Information : MessageBoxImage.Warning;
                MessageBox.Show(
                    dialog,
                    $"{summary}\n客户端：{ip}:{port}\n{pingResult}\n{tcpResult}",
                    "IP测试",
                    MessageBoxButton.OK,
                    icon);
            }
            finally
            {
                testButton.IsEnabled = true;
            }
        };

        dialog.ShowDialog();
    }

    private void ApplyTheme(string theme)
    {
        try
        {
            if (theme == "dark")
            {
                SetBrushColor("WindowBrush", "#1C2430");
                SetBrushColor("TopBarBrush", "#202C3C");
                SetBrushColor("PanelBrush", "#263244");
                SetBrushColor("SurfaceAltBrush", "#2C3A4E");
                SetBrushColor("PanelBorderBrush", "#40506A");
                SetBrushColor("HeadBrush", "#243041");
                SetBrushColor("DividerBrush", "#3C4B63");
                SetBrushColor("InputBrush", "#2D3B4F");
                SetBrushColor("InputBorderBrush", "#51627D");
                SetBrushColor("TitleBrush", "#E8EEF8");
                SetBrushColor("SubtleBrush", "#B4C3D8");
                SetBrushColor("HintBrush", "#93A6BF");
                SetBrushColor("MonitorValueTextBrush", "#F6FAFF");
                SetBrushColor("MonitorValueBackgroundBrush", "#5C4B2A");
                SetBrushColor("MonitorToggleBackgroundBrush", "#32455F");
                SetBrushColor("SelectedRowBrush", "#30455F");
                SetBrushColor("HoverRowBrush", "#344B67");
                SetBrushColor("SuccessBrush", "#4CD17E");
                SetBrushColor("DangerBrush", "#FF7B7B");
                ApplyMenuSystemBrushes();
                return;
            }

            SetBrushColor("WindowBrush", "#EEF1F6");
            SetBrushColor("TopBarBrush", "#F8FAFD");
            SetBrushColor("PanelBrush", "#FFFFFF");
            SetBrushColor("SurfaceAltBrush", "#F7FAFF");
            SetBrushColor("PanelBorderBrush", "#CBD4E1");
            SetBrushColor("HeadBrush", "#E9EFF8");
            SetBrushColor("DividerBrush", "#CAD3E0");
            SetBrushColor("InputBrush", "#FFFFFF");
            SetBrushColor("InputBorderBrush", "#C7D1E0");
            SetBrushColor("TitleBrush", "#132033");
            SetBrushColor("SubtleBrush", "#4B5A70");
            SetBrushColor("HintBrush", "#6A778D");
            SetBrushColor("MonitorValueTextBrush", "#0E2A47");
            SetBrushColor("MonitorValueBackgroundBrush", "#FFF1D6");
            SetBrushColor("MonitorToggleBackgroundBrush", "#F1F8FF");
            SetBrushColor("SelectedRowBrush", "#EAF3FF");
            SetBrushColor("HoverRowBrush", "#F3F8FF");
            SetBrushColor("SuccessBrush", "#1F8A4C");
            SetBrushColor("DangerBrush", "#B42318");
            ApplyMenuSystemBrushes();
        }
        catch (Exception ex)
        {
            AddLog($"主题切换失败: {ex.Message}");
            MessageBox.Show(this, $"主题切换失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ApplyHtmlDialogTheme(Window dialog)
    {
        dialog.Background = GetThemeBrush("PanelBrush", Brushes.White);
        dialog.Foreground = GetThemeBrush("TitleBrush", Brushes.Black);
        dialog.FontFamily = FontFamily;
        TryCopyImplicitStyle(dialog, typeof(TextBlock));
        TryCopyImplicitStyle(dialog, typeof(TextBox));
        TryCopyImplicitStyle(dialog, typeof(Button));
        TryCopyImplicitStyle(dialog, typeof(ComboBox));
        TryCopyImplicitStyle(dialog, typeof(CheckBox));
        TryCopyImplicitStyle(dialog, typeof(RadioButton));
    }

    private void TryCopyImplicitStyle(Window dialog, Type targetType)
    {
        if (Resources[targetType] is Style style)
        {
            dialog.Resources[targetType] = style;
        }
    }

    private void SetBrushColor(string key, string hex)
    {
        if (ColorConverter.ConvertFromString(hex) is not Color color)
        {
            return;
        }

        // Replace brush instance to avoid modifying frozen resources.
        Resources[key] = new SolidColorBrush(color);
    }

    private void ApplyMenuSystemBrushes()
    {
        // Keep WPF system menu popup brushes in sync with app theme resources.
        var panelBrush = GetThemeBrush("PanelBrush", Brushes.White);
        var textBrush = GetThemeBrush("TitleBrush", Brushes.Black);
        var hoverBrush = GetThemeBrush("HoverRowBrush", Brushes.LightBlue);

        SetThemeResource(SystemColors.MenuBrushKey, panelBrush);
        SetThemeResource(SystemColors.MenuTextBrushKey, textBrush);
        SetThemeResource(SystemColors.HighlightBrushKey, hoverBrush);
        SetThemeResource(SystemColors.HighlightTextBrushKey, textBrush);
        SetThemeResource(SystemColors.ControlBrushKey, panelBrush);
        SetThemeResource(SystemColors.ControlTextBrushKey, textBrush);
        SetThemeResource("Menu.Static.Background", panelBrush);
        SetThemeResource("Menu.Static.Foreground", textBrush);
        SetThemeResource("Menu.Static.Border", GetThemeBrush("PanelBorderBrush", Brushes.Gray));
        SetThemeResource("MenuItem.Highlight.Background", hoverBrush);
        SetThemeResource("MenuItem.Highlight.Border", hoverBrush);
        SetThemeResource("MenuItem.Disabled.Foreground", GetThemeBrush("HintBrush", Brushes.Gray));
        SetThemeResource("MenuItem.Selected.Background", GetThemeBrush("SelectedRowBrush", Brushes.SteelBlue));
        SetThemeResource("MenuItem.Selected.Border", GetThemeBrush("PanelBorderBrush", Brushes.Gray));
        SetThemeResource("MenuItem.Submenu.Background", panelBrush);
        SetThemeResource("MenuItem.Submenu.Border", GetThemeBrush("PanelBorderBrush", Brushes.Gray));

        if (panelBrush is SolidColorBrush panelSolid)
        {
            SetThemeResource(SystemColors.MenuColorKey, panelSolid.Color);
            SetThemeResource(SystemColors.ControlColorKey, panelSolid.Color);
        }

        if (textBrush is SolidColorBrush textSolid)
        {
            SetThemeResource(SystemColors.MenuTextColorKey, textSolid.Color);
            SetThemeResource(SystemColors.ControlTextColorKey, textSolid.Color);
        }

        if (hoverBrush is SolidColorBrush hoverSolid)
        {
            SetThemeResource(SystemColors.HighlightColorKey, hoverSolid.Color);
        }
    }

    private void SetThemeResource(object key, object value)
    {
        Resources[key] = value;
        if (Application.Current is not null)
        {
            Application.Current.Resources[key] = value;
        }
    }

    private Brush GetThemeBrush(string key, Brush fallback)
    {
        return Resources[key] as Brush ?? fallback;
    }

    private void RenameNodeButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_selectedNode is null)
        {
            return;
        }

        var oldName = _selectedNode.Name;
        var newName = _selectedNode.NodeType == ProjectNodeType.DataBlock
            ? PromptForDbBlockRename(oldName)
            : PromptForName(oldName);
        if (string.IsNullOrWhiteSpace(newName))
        {
            return;
        }

        newName = newName.Trim();
        if (string.Equals(oldName, newName, StringComparison.Ordinal))
        {
            return;
        }

        _selectedNode.Name = newName;
        OnPropertyChanged(nameof(CurrentDataBlockName));
        AddLog($"已重命名: {oldName} -> {newName}");
    }

    private void DbBlockPropertiesButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_selectedNode is null || _selectedNode.NodeType != ProjectNodeType.DataBlock)
        {
            return;
        }

        var oldName = _selectedNode.Name;
        var newName = PromptForDbBlockProperties(_selectedNode);
        if (string.IsNullOrWhiteSpace(newName) || string.Equals(oldName, newName, StringComparison.Ordinal))
        {
            return;
        }

        if (TryParseDbNumber(oldName, out var oldDbNumber) &&
            TryParseDbNumber(newName, out var newDbNumber) &&
            oldDbNumber != newDbNumber)
        {
            MoveDbDataAndVariables(oldDbNumber, newDbNumber);
        }

        _selectedNode.Name = newName;
        RefreshVariableFilter();
        AddLog($"已更新数据块属性: {oldName} -> {newName}");
    }

    private void AddDbBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        var parent = ResolveDataBlockParent();
        var nextDb = FindNextDbNumber();
        var name = PromptForDbBlockName(parent, nextDb);
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        AddDbBlock(name, parent);
    }

    private void RemoveDbBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_selectedNode is null || _selectedNode.NodeType != ProjectNodeType.DataBlock)
        {
            AddLog("请先在项目树中选中一个数据块后再删除。");
            return;
        }

        _selectedNode.Parent?.Children.Remove(_selectedNode);
        AddLog($"已删除数据块: {_selectedNode.Name}");
        RefreshVariableFilter();
    }

    private void AddDbBlock(string name, ProjectTreeNode? parent = null)
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            throw new InvalidOperationException("未找到可用 PLC 节点。");
        }

        var targetParent = parent ?? GetFolderNode(plcRoot, ProjectNodeType.ProgramFolder);
        var dbNumber = ParseDbNumber(name);
        if (IsDbNumberInUse(targetParent, dbNumber))
        {
            throw new InvalidOperationException($"DB{dbNumber} 已存在，请更换编号。");
        }

        var node = new ProjectTreeNode
        {
            Name = name,
            NodeType = ProjectNodeType.DataBlock,
            Parent = targetParent
        };
        targetParent.Children.Add(node);

        EnsureDbMemoryReady(dbNumber);
        AddLog($"已新增数据块: {name}");
    }

    private int FindNextDbNumber()
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            throw new InvalidOperationException("未找到可用 PLC 节点。");
        }

        var dbFolder = ResolveDataBlockParent();
        return FindNextDbNumber(dbFolder);
    }

    private int FindNextDbNumber(ProjectTreeNode dbFolder, int? excludeDbNumber = null)
    {
        var usedNumbers = new HashSet<int>();
        foreach (var node in dbFolder.Children)
        {
            if (node.NodeType != ProjectNodeType.DataBlock)
            {
                continue;
            }

            if (TryParseDbNumber(node.Name, out var dbNumber))
            {
                if (excludeDbNumber.HasValue && dbNumber == excludeDbNumber.Value)
                {
                    continue;
                }
                usedNumbers.Add(dbNumber);
            }
        }

        var candidate = 1;
        while (usedNumbers.Contains(candidate))
        {
            candidate++;
        }

        return candidate;
    }

    private void AddPlcSimulator(string plcName)
    {
        var plcRoot = new ProjectTreeNode
        {
            Name = plcName,
            NodeType = ProjectNodeType.PlcRoot
        };

        var programFolder = new ProjectTreeNode
        {
            Name = "程序块",
            NodeType = ProjectNodeType.ProgramFolder,
            Parent = plcRoot
        };
        plcRoot.Children.Add(programFolder);
        ProjectNodes.Add(plcRoot);
        SortPlcRootsByIndex();

        _selectedNode = plcRoot;
        var defaultAddress = ProjectNodes.Count == 1 ? GetPreferredDefaultAddress() : string.Empty;
        var settings = new PlcEndpointSettings(defaultAddress, DefaultPlcPort, true);
        _plcEndpointSettings[plcRoot] = settings;
        ApplyPlcSettings(plcRoot, settings);
        AddLog($"已新增 PLC 模拟器: {plcName}");
        RefreshVariableFilter();
        UpdateSimulatorRunState();
    }

    private int FindNextPlcIndex()
    {
        var used = new HashSet<int>();
        foreach (var node in ProjectNodes.Where(x => x.NodeType == ProjectNodeType.PlcRoot))
        {
            var num = ParsePlcIndexFromName(node.Name);
            if (num > 0)
            {
                used.Add(num);
            }
        }

        var candidate = 1;
        while (used.Contains(candidate))
        {
            candidate++;
        }

        return candidate;
    }

    private void SortPlcRootsByIndex()
    {
        var sorted = ProjectNodes
            .OrderBy(x => x.NodeType == ProjectNodeType.PlcRoot ? 0 : 1)
            .ThenBy(x =>
            {
                var idx = ParsePlcIndexFromName(x.Name);
                return idx > 0 ? idx : int.MaxValue;
            })
            .ThenBy(x => x.Name, StringComparer.Ordinal)
            .ToList();

        if (sorted.SequenceEqual(ProjectNodes))
        {
            return;
        }

        ProjectNodes.Clear();
        foreach (var node in sorted)
        {
            ProjectNodes.Add(node);
        }
    }

    private static int ParsePlcIndexFromName(string name)
    {
        const string marker = "模拟器_";
        var idx = name.LastIndexOf(marker, StringComparison.Ordinal);
        if (idx < 0)
        {
            return -1;
        }

        var start = idx + marker.Length;
        return start < name.Length && int.TryParse(name[start..], out var num) ? num : -1;
    }

    private ProjectTreeNode? GetCurrentPlcRoot()
    {
        if (_selectedNode is null)
        {
            return ProjectNodes.FirstOrDefault();
        }

        var cursor = _selectedNode;
        while (cursor is not null && cursor.NodeType != ProjectNodeType.PlcRoot)
        {
            cursor = cursor.Parent;
        }

        return cursor ?? ProjectNodes.FirstOrDefault();
    }

    private static ProjectTreeNode GetFolderNode(ProjectTreeNode plcRoot, ProjectNodeType folderType)
    {
        var folder = FindNodeByType(plcRoot, folderType);
        if (folder is null)
        {
            throw new InvalidOperationException($"PLC 鑺傜偣缂哄皯 {folderType} 瀛愮洰褰曘€");
        }

        return folder;
    }

    private ProjectTreeNode ResolveDataBlockParent()
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            throw new InvalidOperationException("未找到可用 PLC 节点。");
        }

        return GetFolderNode(plcRoot, ProjectNodeType.ProgramFolder);
    }

    private int? ResolveSelectedDbNumber()
    {
        if (_selectedNode is null)
        {
            return null;
        }

        if (_selectedNode.NodeType == ProjectNodeType.DataBlock)
        {
            if (TryParseDbNumber(_selectedNode.Name, out var dbNumber))
            {
                return dbNumber;
            }

            AddLog($"数据块名称无法解析编号：{_selectedNode.Name}。请使用 DBx 或 名称 [DBx] 格式。");
            return null;
        }

        return null;
    }

    private void RefreshVariableFilter()
    {
        _activeDbNumber = ResolveSelectedDbNumber();
        EnsureVariablesSortOrder();
        RefreshFilteredVariablesSafely();
        OnPropertyChanged(nameof(CurrentDataBlockName));
        OnPropertyChanged(nameof(IsPlcOverviewSelected));
        OnPropertyChanged(nameof(IsProgramAreaSelected));
        OnPropertyChanged(nameof(IsDataBlockSelected));
        OnPropertyChanged(nameof(IsVariablePanelVisible));
    }

    private int ParseDbNumber(string dbDisplayName)
    {
        if (!TryParseDbNumber(dbDisplayName, out var dbNumber))
        {
            throw new InvalidOperationException($"无法识别数据块编号: {dbDisplayName}");
        }

        return dbNumber;
    }

    private bool IsDbNumberInUse(ProjectTreeNode dbFolder, int dbNumber, ProjectTreeNode? excludeNode = null)
    {
        foreach (var node in dbFolder.Children)
        {
            if (node.NodeType != ProjectNodeType.DataBlock)
            {
                continue;
            }

            if (excludeNode is not null && ReferenceEquals(node, excludeNode))
            {
                continue;
            }

            if (TryParseDbNumber(node.Name, out var existingDbNumber) && existingDbNumber == dbNumber)
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseDbNumber(string dbDisplayName, out int dbNumber)
    {
        return TryParseDbDisplayName(dbDisplayName, out _, out dbNumber);
    }

    private static bool TryParseDbDisplayName(string dbDisplayName, out string customName, out int dbNumber)
    {
        customName = string.Empty;
        dbNumber = 0;
        if (string.IsNullOrWhiteSpace(dbDisplayName))
        {
            return false;
        }

        var markerStart = dbDisplayName.LastIndexOf("[DB", StringComparison.OrdinalIgnoreCase);
        if (markerStart >= 0)
        {
            var numberStart = markerStart + 3;
            var markerEnd = dbDisplayName.IndexOf(']', numberStart);
            if (markerEnd > numberStart &&
                int.TryParse(dbDisplayName[numberStart..markerEnd], out dbNumber) &&
                dbNumber > 0)
            {
                customName = dbDisplayName[..markerStart].Trim();
                if (string.IsNullOrWhiteSpace(customName))
                {
                    customName = $"数据块_{dbNumber}";
                }
                return true;
            }
        }

        var openIdx = dbDisplayName.IndexOf(' ', StringComparison.Ordinal);
        var dbName = openIdx > 0 ? dbDisplayName[..openIdx] : dbDisplayName;
        if (dbName.StartsWith("DB", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(dbName[2..], out dbNumber) &&
            dbNumber > 0)
        {
            customName = openIdx > 0 ? dbDisplayName[(openIdx + 1)..].Trim() : $"数据块_{dbNumber}";
            if (string.IsNullOrWhiteSpace(customName))
            {
                customName = $"数据块_{dbNumber}";
            }

            return true;
        }

        return false;
    }

    private static string BuildDbDisplayName(int dbNumber, string customName)
    {
        return $"{customName} [DB{dbNumber}]";
    }

    private void EnsureDbMemoryReady(int dbNumber)
    {
        var _ = _memoryStore.EnsureDb(dbNumber, DefaultDbSize, out var created);
        if (created)
        {
            AddLog($"已创建 DB{dbNumber} 独立内存区，大小={DefaultDbSize} 字节");
        }

        foreach (var host in _plcServerHosts.Values)
        {
            host.EnsureDbAreaRegistered(dbNumber);
        }
    }

    private void MoveDbDataAndVariables(int sourceDbNumber, int targetDbNumber)
    {
        if (sourceDbNumber == targetDbNumber)
        {
            return;
        }

        if (_memoryStore.TryGetDb(sourceDbNumber, out var sourceDb) && sourceDb is not null)
        {
            var targetDb = _memoryStore.EnsureDb(targetDbNumber, sourceDb.Length, out _);
            Buffer.BlockCopy(sourceDb, 0, targetDb, 0, sourceDb.Length);
        }
        else
        {
            EnsureDbMemoryReady(targetDbNumber);
        }

        var movedRows = Variables.Where(x => x.DbNumber == sourceDbNumber).ToList();
        foreach (var row in movedRows)
        {
            row.DbNumber = targetDbNumber;
        }

        RecalculateOffsetsForDb(sourceDbNumber);
        RecalculateOffsetsForDb(targetDbNumber);
    }

    private static ProjectTreeNode? FindNodeByType(ProjectTreeNode root, ProjectNodeType nodeType)
    {
        if (root.NodeType == nodeType)
        {
            return root;
        }

        foreach (var child in root.Children)
        {
            var found = FindNodeByType(child, nodeType);
            if (found is not null)
            {
                return found;
            }
        }

        return null;
    }

    private void StartButton_OnClick(object sender, RoutedEventArgs e)
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            return;
        }

        try
        {
            if (!int.TryParse(PortTextBox.Text.Trim(), out var port))
            {
                throw new InvalidOperationException("端口格式无效。");
            }
            if (port is < 1 or > 65535)
            {
                throw new InvalidOperationException("端口范围必须在 1-65535。");
            }

            var address = AddressTextBox.Text.Trim();
            var settings = new PlcEndpointSettings(address, port, true);
            _plcEndpointSettings[plcRoot] = settings;
            ApplyPlcSettings(plcRoot, settings);
            if (port <= 1024)
            {
                AddLog("提示: 当前使用特权端口，Windows 下通常需要管理员权限。");
            }
        }
        catch (Exception ex)
        {
            AddLog($"启动失败: {ex.Message}");
            UpdateConnectionStatus(false);
            MessageBox.Show($"启动失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void StopButton_OnClick(object sender, RoutedEventArgs e)
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            return;
        }

        StopPlcServer(plcRoot);
        UpdateConnectionStatus(_plcServerHosts.Count > 0);
    }

    private void MonitorToggleButton_OnClick(object sender, RoutedEventArgs e)
    {
        SetMonitorEnabled(!IsMonitorEnabled);
    }

    private void AddRowButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var selectedDbNumber = ResolveSelectedDbNumber();
            if (!selectedDbNumber.HasValue)
            {
                AddLog("请先在项目树中选中一个数据块，再新增变量。");
                return;
            }

            var row = new VariableRow
            {
                Name = $"Var{Variables.Count + 1}",
                DbNumber = selectedDbNumber.Value,
                Offset = 0,
                Bit = 0,
                Type = PlcValueType.Bool,
                StringLength = DefaultStringLength,
                StartValue = GetDefaultStartValueByType(PlcValueType.Bool),
                Value = "false"
            };

            Variables.Add(row);
            RecalculateOffsetsForDb(row.DbNumber);
            TryWriteRow(row, writeLog: false);
        }
        catch (Exception ex)
        {
            AddLog($"新增变量失败: {ex.Message}");
            MessageBox.Show(this, $"新增变量失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DeleteRowButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (VariablesGrid.SelectedItem is not VariableRow row)
        {
            return;
        }

        TryClearRowValue(row);
        Variables.Remove(row);
        RecalculateOffsetsForDb(row.DbNumber);
    }

    private void VariablesGrid_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isMonitorValueCellContext = false;
        var source = e.OriginalSource as DependencyObject;
        var cell = FindVisualParent<DataGridCell>(source);
        if (cell is null)
        {
            return;
        }

        var row = FindVisualParent<DataGridRow>(cell);
        if (row?.Item is VariableRow variable)
        {
            VariablesGrid.SelectedItem = variable;
            VariablesGrid.CurrentCell = new DataGridCellInfo(variable, cell.Column);
            _isMonitorValueCellContext = ReferenceEquals(cell.Column, MonitorValueColumn);
        }
    }

    private void VariablesGrid_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        var selectedRow = VariablesGrid.SelectedItem as VariableRow;
        var canToggleArrayExpand = selectedRow?.Type == PlcValueType.Array;
        ToggleArrayExpandMenuItem.Visibility = canToggleArrayExpand ? Visibility.Visible : Visibility.Collapsed;
        ArrayExpandSeparator.Visibility = canToggleArrayExpand ? Visibility.Visible : Visibility.Collapsed;
        if (canToggleArrayExpand)
        {
            ToggleArrayExpandMenuItem.Header = selectedRow!.IsArrayExpanded ? "收起数组" : "展开数组";
        }

        var isMonitorCell = _isMonitorValueCellContext || ReferenceEquals(VariablesGrid.CurrentCell.Column, MonitorValueColumn);
        var canEditMonitorValue = IsMonitorEnabled && isMonitorCell && VariablesGrid.SelectedItem is VariableRow;
        EditMonitorValueMenuItem.Visibility = canEditMonitorValue ? Visibility.Visible : Visibility.Collapsed;
        MonitorEditSeparator.Visibility = canEditMonitorValue ? Visibility.Visible : Visibility.Collapsed;
    }

    private void EditMonitorValueMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        if (VariablesGrid.SelectedItem is not VariableRow row)
        {
            return;
        }

        var input = PromptForValue(row);
        if (input is null)
        {
            return;
        }

        var oldStartValue = row.StartValue;
        row.StartValue = input;
        try
        {
            ValidateRowAddress(row);
            WriteRowValue(row);
            row.Value = ReadRowValue(row);
            AddLog($"右键写入 {BuildAddressText(row)} = {row.StartValue}");
        }
        catch (Exception ex)
        {
            row.StartValue = oldStartValue;
            AddLog($"右键写入失败 {BuildAddressText(row)}: {ex.Message}");
            MessageBox.Show($"写入失败: {ex.Message}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void ToggleArrayExpandMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        if (VariablesGrid.SelectedItem is not VariableRow row || row.Type != PlcValueType.Array)
        {
            return;
        }

        row.IsArrayExpanded = !row.IsArrayExpanded;
        VariablesGrid.Items.Refresh();
    }

    private void TryReadRow(VariableRow row, bool writeLog = true)
    {
        try
        {
            ValidateRowAddress(row);
            row.Value = ReadRowValue(row);
            if (writeLog)
            {
                AddLog($"读取 {BuildAddressText(row)} = {row.Value}");
            }
        }
        catch (Exception ex)
        {
            if (writeLog)
            {
                AddLog($"读取失败 {BuildAddressText(row)}: {ex.Message}");
            }
        }
    }

    private void TryWriteRow(VariableRow row, bool writeLog = true)
    {
        try
        {
            ValidateRowAddress(row);
            WriteRowValue(row);
            if (writeLog)
            {
                AddLog($"写入 {BuildAddressText(row)} = {row.StartValue}");
            }
        }
        catch (Exception ex)
        {
            if (writeLog)
            {
                AddLog($"写入失败 {BuildAddressText(row)}: {ex.Message}");
            }
        }
    }

    private void VariablesGrid_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_isRefreshingValues || e.EditAction != DataGridEditAction.Commit)
        {
            return;
        }

        if (e.Row.Item is not VariableRow row)
        {
            return;
        }

        if (e.Column == StartValueColumn)
        {
            if (row.Type == PlcValueType.Array)
            {
                return;
            }

            Dispatcher.InvokeAsync(() => TryWriteRow(row));
            return;
        }

        if (e.Column == TypeColumn)
        {
            Dispatcher.InvokeAsync(() =>
            {
                if (row.Type == PlcValueType.Array)
                {
                    var definition = PromptForArrayTypeDefinition(
                        row.ArrayElementType,
                        row.ArrayLowerBound,
                        row.ArrayUpperBound,
                        row.ArrayElementType == PlcValueType.String ? row.StringLength : null,
                        allowNegativeIndex: false);
                    if (definition is null)
                    {
                        if (_typeEditorTextBeforeEdit.TryGetValue(row, out var oldTypeEditorText))
                        {
                            row.TypeEditorText = oldTypeEditorText;
                        }
                    }
                    else
                    {
                        row.ArrayLowerBound = definition.LowerBound;
                        row.ArrayUpperBound = definition.UpperBound;
                        row.ArrayElementType = definition.ElementType;
                        if (definition.StringLength is int len)
                        {
                            row.StringLength = len;
                        }

                        row.TypeEditorText = BuildArrayTypeEditorText(definition);
                        row.StartValue = BuildDefaultArrayStartValue(row);
                        row.IsArrayExpanded = true;
                    }
                }

                RecalculateOffsetsForDb(row.DbNumber);
                TryWriteRow(row, writeLog: false);
                _typeEditorTextBeforeEdit.Remove(row);

                // Force row-details refresh so newly created array rows expand immediately.
                Dispatcher.InvokeAsync(() =>
                {
                    VariablesGrid.Items.Refresh();
                    if (row.Type == PlcValueType.Array &&
                        VariablesGrid.ItemContainerGenerator.ContainerFromItem(row) is DataGridRow rowContainer)
                    {
                        rowContainer.DetailsVisibility = Visibility.Visible;
                    }
                }, DispatcherPriority.Background);
            });
        }
    }

    private void VariablesGrid_OnBeginningEdit(object sender, DataGridBeginningEditEventArgs e)
    {
        if (e.Row.Item is not VariableRow row)
        {
            return;
        }

        if (e.Column == StartValueColumn && row.Type == PlcValueType.Array)
        {
            e.Cancel = true;
            return;
        }

        if (e.Column != TypeColumn)
        {
            return;
        }

        _typeEditorTextBeforeEdit[row] = row.TypeEditorText;
    }

    private void Variables_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<VariableRow>())
            {
                item.PropertyChanged += VariableRow_OnPropertyChanged;
                RecalculateOffsetsForDb(item.DbNumber);
            }
        }

        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<VariableRow>())
            {
                item.PropertyChanged -= VariableRow_OnPropertyChanged;
                RecalculateOffsetsForDb(item.DbNumber);
            }
        }
    }

    private void VariableRow_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not VariableRow row)
        {
            return;
        }

        if (e.PropertyName == nameof(VariableRow.Type))
        {
            row.StartValue = row.Type == PlcValueType.Array
                ? BuildDefaultArrayStartValue(row)
                : GetDefaultStartValueByType(row.Type);
            RecalculateOffsetsForDb(row.DbNumber);
            TryWriteRow(row, writeLog: false);
            return;
        }

        if (e.PropertyName == nameof(VariableRow.StartValue) && row.Type == PlcValueType.Array && !_isRefreshingValues)
        {
            TryWriteRow(row, writeLog: false);
        }
    }

    private void ValueRefreshTimer_OnTick(object? sender, EventArgs e)
    {
        if (_isRefreshingValues)
        {
            return;
        }

        _isRefreshingValues = true;
        try
        {
            var targetRows = _activeDbNumber.HasValue
                ? Variables.Where(x => x.DbNumber == _activeDbNumber.Value).ToList()
                : Variables.ToList();
            foreach (var row in targetRows)
            {
                TryReadRow(row, writeLog: false);
            }
        }
        finally
        {
            _isRefreshingValues = false;
        }
    }

    private string ReadRowValue(VariableRow row)
    {
        return row.Type switch
        {
            PlcValueType.Bool => _memoryStore.GetBool(row.DbNumber, row.Offset, row.Bit).ToString(),
            PlcValueType.Array => ReadArrayValue(row),
            PlcValueType.Byte => _memoryStore.Read(row.DbNumber, row.Offset, 1)[0].ToString(CultureInfo.InvariantCulture),
            PlcValueType.Usint => _memoryStore.Read(row.DbNumber, row.Offset, 1)[0].ToString(CultureInfo.InvariantCulture),
            PlcValueType.Sint => unchecked((sbyte)_memoryStore.Read(row.DbNumber, row.Offset, 1)[0]).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Word => ReadUInt16(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Uint => ReadUInt16(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Dword => ReadUInt32(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Udint => ReadUInt32(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Char => Encoding.ASCII.GetString(_memoryStore.Read(row.DbNumber, row.Offset, 1)),
            PlcValueType.Int => _memoryStore.GetInt(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Dint => _memoryStore.GetDInt(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Real => _memoryStore.GetReal(row.DbNumber, row.Offset).ToString("0.###", CultureInfo.InvariantCulture),
            PlcValueType.LReal => ReadLReal(row.DbNumber, row.Offset).ToString("0.######", CultureInfo.InvariantCulture),
            PlcValueType.S5Time => DecodeS5TimeToMs(ReadUInt16(row.DbNumber, row.Offset)).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Time => _memoryStore.GetDInt(row.DbNumber, row.Offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Date => S7DateBase.AddDays(ReadUInt16(row.DbNumber, row.Offset)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            PlcValueType.TimeOfDay => FormatTimeOfDay(ReadUInt32(row.DbNumber, row.Offset)),
            PlcValueType.DateAndTime => DecodeDateAndTime(_memoryStore.Read(row.DbNumber, row.Offset, 8)),
            PlcValueType.String => GetS7String(row.DbNumber, row.Offset),
            PlcValueType.WString => GetWString(row.DbNumber, row.Offset),
            _ => throw new NotSupportedException($"不支持的类型: {row.Type}")
        };
    }

    private void WriteRowValue(VariableRow row)
    {
        switch (row.Type)
        {
            case PlcValueType.Bool:
                _memoryStore.SetBool(row.DbNumber, row.Offset, row.Bit, ParseBool(row.StartValue));
                return;
            case PlcValueType.Array:
                WriteArrayValue(row);
                return;
            case PlcValueType.Byte:
                _memoryStore.Write(row.DbNumber, row.Offset, [byte.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture)]);
                return;
            case PlcValueType.Usint:
                _memoryStore.Write(row.DbNumber, row.Offset, [byte.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture)]);
                return;
            case PlcValueType.Sint:
                _memoryStore.Write(row.DbNumber, row.Offset, [unchecked((byte)sbyte.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture))]);
                return;
            case PlcValueType.Word:
            case PlcValueType.Uint:
                WriteUInt16(row.DbNumber, row.Offset, ushort.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Dword:
            case PlcValueType.Udint:
                WriteUInt32(row.DbNumber, row.Offset, uint.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Char:
                _memoryStore.Write(row.DbNumber, row.Offset, [ParseChar(row.StartValue)]);
                return;
            case PlcValueType.Int:
                _memoryStore.SetInt(row.DbNumber, row.Offset, short.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Dint:
                _memoryStore.SetDInt(row.DbNumber, row.Offset, int.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Real:
                _memoryStore.SetReal(row.DbNumber, row.Offset, float.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.LReal:
                WriteLReal(row.DbNumber, row.Offset, double.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.S5Time:
                WriteUInt16(row.DbNumber, row.Offset, EncodeMsToS5Time(ulong.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture)));
                return;
            case PlcValueType.Time:
                _memoryStore.SetDInt(row.DbNumber, row.Offset, int.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Date:
                var date = DateOnly.Parse(row.StartValue.Trim(), CultureInfo.InvariantCulture);
                WriteUInt16(row.DbNumber, row.Offset, checked((ushort)(date.DayNumber - S7DateBase.DayNumber)));
                return;
            case PlcValueType.TimeOfDay:
                WriteUInt32(row.DbNumber, row.Offset, ParseTimeOfDay(row.StartValue));
                return;
            case PlcValueType.DateAndTime:
                _memoryStore.Write(row.DbNumber, row.Offset, EncodeDateAndTime(row.StartValue));
                return;
            case PlcValueType.String:
                SetS7String(row.DbNumber, row.Offset, row.StringLength, NormalizeStringStartValue(row.StartValue));
                return;
            case PlcValueType.WString:
                SetWString(row.DbNumber, row.Offset, row.StringLength, NormalizeStringStartValue(row.StartValue));
                return;
            default:
                throw new NotSupportedException($"不支持的类型: {row.Type}");
        }
    }

    private void TryClearRowValue(VariableRow row)
    {
        try
        {
            ValidateRowAddress(row);
            switch (row.Type)
            {
                case PlcValueType.Bool:
                    _memoryStore.SetBool(row.DbNumber, row.Offset, row.Bit, false);
                    break;
                case PlcValueType.Array:
                    ClearArrayValue(row);
                    break;
                case PlcValueType.Byte:
                case PlcValueType.Usint:
                    _memoryStore.Write(row.DbNumber, row.Offset, [0]);
                    break;
                case PlcValueType.Sint:
                    _memoryStore.Write(row.DbNumber, row.Offset, [0]);
                    break;
                case PlcValueType.Word:
                case PlcValueType.Uint:
                    WriteUInt16(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Dword:
                case PlcValueType.Udint:
                    WriteUInt32(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Char:
                    _memoryStore.Write(row.DbNumber, row.Offset, [0]);
                    break;
                case PlcValueType.Int:
                    _memoryStore.SetInt(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Dint:
                    _memoryStore.SetDInt(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Real:
                    _memoryStore.SetReal(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.LReal:
                    WriteLReal(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.S5Time:
                    WriteUInt16(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Time:
                    _memoryStore.SetDInt(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.Date:
                    WriteUInt16(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.TimeOfDay:
                    WriteUInt32(row.DbNumber, row.Offset, 0);
                    break;
                case PlcValueType.DateAndTime:
                    _memoryStore.Write(row.DbNumber, row.Offset, new byte[8]);
                    break;
                case PlcValueType.String:
                    SetS7String(row.DbNumber, row.Offset, row.StringLength, string.Empty);
                    break;
                case PlcValueType.WString:
                    SetWString(row.DbNumber, row.Offset, row.StringLength, string.Empty);
                    break;
                default:
                    throw new NotSupportedException($"不支持的类型: {row.Type}");
            }

            AddLog($"已清空 {BuildAddressText(row)}");
        }
        catch (Exception ex)
        {
            AddLog($"清空失败 {BuildAddressText(row)}: {ex.Message}");
        }
    }

    private string BuildAddressText(VariableRow row)
    {
        return row.AddressDisplay;
    }

    private void AddLog(string message)
    {
        Dispatcher.InvokeAsync(() =>
        {
            var text = $"[{DateTime.Now:HH:mm:ss}] {message}";
            LogListBox.Items.Insert(0, text);
            if (LogListBox.Items.Count > 500)
            {
                LogListBox.Items.RemoveAt(LogListBox.Items.Count - 1);
            }

            UpdateLogCountText();
        });
    }

    private void ClearLogButton_OnClick(object sender, RoutedEventArgs e)
    {
        LogListBox.Items.Clear();
        UpdateLogCountText();
    }

    private void UpdateLogCountText()
    {
        LogCountTextBlock.Text = $"最近 {LogListBox.Items.Count} 条";
    }

    private void StopServer()
    {
        var plcRoot = GetCurrentPlcRoot();
        if (plcRoot is null)
        {
            return;
        }

        StopPlcServer(plcRoot);
        UpdateConnectionStatus(_plcServerHosts.Count > 0);
    }

    private void SetMonitorEnabled(bool enabled, bool writeLog = true)
    {
        IsMonitorEnabled = enabled;
        MonitorValueColumn.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        MonitorValueColumn.DisplayIndex = 4;
        AddressColumn.DisplayIndex = enabled ? 5 : 4;
        ApplyEqualColumnWidths();

        MonitorTogglePanel.Background = enabled
            ? GetThemeBrush("MonitorToggleBackgroundBrush", Brushes.Transparent)
            : Brushes.Transparent;
        MonitorToggleButton.Content = enabled ? "⏸" : "▶";
        MonitorToggleButton.ToolTip = enabled ? "停止监控（隐藏监视值列）" : "启动监控（显示监视值列）";

        if (enabled)
        {
            if (!_valueRefreshTimer.IsEnabled)
            {
                _valueRefreshTimer.Start();
            }

            ValueRefreshTimer_OnTick(this, EventArgs.Empty);
        }
        else
        {
            _valueRefreshTimer.Stop();
        }

        if (writeLog)
        {
            AddLog(enabled ? "已启动变量监控。" : "已停止变量监控。");
        }
    }

    private void ApplyEqualColumnWidths()
    {
        NameColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        TypeColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        OffsetColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        StartValueColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        MonitorValueColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        AddressColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
    }

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void UpdateConnectionStatus(bool connected)
    {
        Title = connected ? "S7 PLC Simulator (已连接)" : "S7 PLC Simulator";
        UpdateSimulatorRunState();
    }

    private void UpdateSimulatorRunState()
    {
        var plcRoots = ProjectNodes.Where(x => x.NodeType == ProjectNodeType.PlcRoot).ToList();

        SimulatorRunStateLampPanel.Children.Clear();
        if (plcRoots.Count == 0)
        {
            return;
        }

        const int maxLampsPerRow = 10;
        const double lampSpacing = 6;
        var lampsPerRow = Math.Min(maxLampsPerRow, plcRoots.Count);
        var rowContainer = new Grid
        {
            Width = lampsPerRow * (10 + lampSpacing),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var rowCount = (int)Math.Ceiling(plcRoots.Count / (double)lampsPerRow);
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            rowContainer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var rowPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 2, 0, 2)
            };
            Grid.SetRow(rowPanel, rowIndex);
            rowContainer.Children.Add(rowPanel);
        }

        for (var i = 0; i < plcRoots.Count; i++)
        {
            var plcRoot = plcRoots[i];
            var plcRunning = _plcServerHosts.ContainsKey(plcRoot);
            var lamp = new Ellipse
            {
                Width = 10,
                Height = 10,
                Fill = plcRunning
                    ? GetThemeBrush("SuccessBrush", Brushes.ForestGreen)
                    : GetThemeBrush("DangerBrush", Brushes.DarkRed),
                Margin = new Thickness(3, 0, 3, 0),
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = $"{plcRoot.Name}: {(plcRunning ? "运行中" : "未运行")}"
            };

            var rowIndex = i / lampsPerRow;
            if (rowContainer.Children[rowIndex] is StackPanel rowPanel)
            {
                rowPanel.Children.Add(lamp);
            }
        }

        SimulatorRunStateLampPanel.Children.Add(rowContainer);
    }

    private static void ValidateRowAddress(VariableRow row)
    {
        if (row.DbNumber <= 0)
        {
            throw new InvalidOperationException($"DB 鍙峰繀椤诲ぇ浜?0锛屽綋鍓嶅€? {row.DbNumber}");
        }

        if (row.Offset < 0)
        {
            throw new InvalidOperationException($"鍋忕Щ閲忎笉鑳戒负璐熸暟锛屽綋鍓嶅€? {row.Offset}");
        }

        if (row.Type == PlcValueType.Bool && (row.Bit < 0 || row.Bit > 7))
        {
            throw new InvalidOperationException($"BOOL 浣嶇储寮曞繀椤诲湪 0-7锛屽綋鍓嶅€? {row.Bit}");
        }

    }

    private static bool ParseBool(string rawValue)
    {
        var value = rawValue.Trim();
        if (bool.TryParse(value, out var boolValue))
        {
            return boolValue;
        }

        return value switch
        {
            "1" => true,
            "0" => false,
            _ => throw new FormatException($"BOOL 鍊兼棤鏁? \"{rawValue}\"锛岃浣跨敤 true/false 鎴?1/0")
        };
    }

    private static string GetDefaultStartValueByType(PlcValueType type)
    {
        return type switch
        {
            PlcValueType.Bool => "FALSE",
            PlcValueType.String => "''",
            PlcValueType.Date => "1990-01-01",
            PlcValueType.TimeOfDay => "00:00:00.000",
            PlcValueType.DateAndTime => "1990-01-01 00:00:00.000",
            _ => "0"
        };
    }

    private static string BuildDefaultArrayStartValue(VariableRow row)
    {
        var elementDefault = GetDefaultStartValueByType(row.ArrayElementType);
        return $"[{string.Join(", ", Enumerable.Repeat(elementDefault, row.ArrayLength))}]";
    }

    private string ReadArrayValue(VariableRow row)
    {
        var (elementSize, _) = GetScalarTypeLayout(row.ArrayElementType, row.StringLength);
        var values = new string[row.ArrayLength];
        for (var i = 0; i < row.ArrayLength; i++)
        {
            var elementOffset = row.Offset + i * elementSize;
            values[i] = ReadArrayElementValue(row.ArrayElementType, row.DbNumber, elementOffset, row.StringLength);
        }

        return $"[{string.Join(", ", values)}]";
    }

    private void WriteArrayValue(VariableRow row)
    {
        var values = ParseArrayValues(row.StartValue, row.ArrayLength);
        var (elementSize, _) = GetScalarTypeLayout(row.ArrayElementType, row.StringLength);
        for (var i = 0; i < values.Length; i++)
        {
            var elementOffset = row.Offset + i * elementSize;
            WriteArrayElementValue(row.ArrayElementType, row.DbNumber, elementOffset, values[i], row.StringLength);
        }
    }

    private void ClearArrayValue(VariableRow row)
    {
        var defaultValue = GetDefaultStartValueByType(row.ArrayElementType);
        var (elementSize, _) = GetScalarTypeLayout(row.ArrayElementType, row.StringLength);
        for (var i = 0; i < row.ArrayLength; i++)
        {
            var elementOffset = row.Offset + i * elementSize;
            WriteArrayElementValue(row.ArrayElementType, row.DbNumber, elementOffset, defaultValue, row.StringLength);
        }
    }

    private static string[] ParseArrayValues(string rawValue, int expectedCount)
    {
        var value = rawValue.Trim();
        if (value.StartsWith('[') && value.EndsWith(']'))
        {
            value = value[1..^1].Trim();
        }

        var items = string.IsNullOrWhiteSpace(value)
            ? []
            : value.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        if (items.Length != expectedCount)
        {
            throw new FormatException($"ARRAY 元素数量不匹配，期望 {expectedCount}，实际 {items.Length}。");
        }

        return items;
    }

    private string ReadArrayElementValue(PlcValueType elementType, int dbNumber, int offset, int stringLength)
    {
        return elementType switch
        {
            PlcValueType.Bool => (_memoryStore.Read(dbNumber, offset, 1)[0] != 0).ToString(),
            PlcValueType.Byte => _memoryStore.Read(dbNumber, offset, 1)[0].ToString(CultureInfo.InvariantCulture),
            PlcValueType.Usint => _memoryStore.Read(dbNumber, offset, 1)[0].ToString(CultureInfo.InvariantCulture),
            PlcValueType.Sint => unchecked((sbyte)_memoryStore.Read(dbNumber, offset, 1)[0]).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Word => ReadUInt16(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Uint => ReadUInt16(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Dword => ReadUInt32(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Udint => ReadUInt32(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Char => Encoding.ASCII.GetString(_memoryStore.Read(dbNumber, offset, 1)),
            PlcValueType.Int => _memoryStore.GetInt(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Dint => _memoryStore.GetDInt(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Real => _memoryStore.GetReal(dbNumber, offset).ToString("0.###", CultureInfo.InvariantCulture),
            PlcValueType.LReal => ReadLReal(dbNumber, offset).ToString("0.######", CultureInfo.InvariantCulture),
            PlcValueType.S5Time => DecodeS5TimeToMs(ReadUInt16(dbNumber, offset)).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Time => _memoryStore.GetDInt(dbNumber, offset).ToString(CultureInfo.InvariantCulture),
            PlcValueType.Date => S7DateBase.AddDays(ReadUInt16(dbNumber, offset)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            PlcValueType.TimeOfDay => FormatTimeOfDay(ReadUInt32(dbNumber, offset)),
            PlcValueType.DateAndTime => DecodeDateAndTime(_memoryStore.Read(dbNumber, offset, 8)),
            PlcValueType.String => GetS7String(dbNumber, offset),
            _ => throw new NotSupportedException($"不支持的数组元素类型: {elementType}")
        };
    }

    private void WriteArrayElementValue(PlcValueType elementType, int dbNumber, int offset, string value, int stringLength)
    {
        switch (elementType)
        {
            case PlcValueType.Bool:
                _memoryStore.Write(dbNumber, offset, [ParseBool(value) ? (byte)1 : (byte)0]);
                return;
            case PlcValueType.Byte:
            case PlcValueType.Usint:
                _memoryStore.Write(dbNumber, offset, [byte.Parse(value.Trim(), CultureInfo.InvariantCulture)]);
                return;
            case PlcValueType.Sint:
                _memoryStore.Write(dbNumber, offset, [unchecked((byte)sbyte.Parse(value.Trim(), CultureInfo.InvariantCulture))]);
                return;
            case PlcValueType.Word:
            case PlcValueType.Uint:
                WriteUInt16(dbNumber, offset, ushort.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Dword:
            case PlcValueType.Udint:
                WriteUInt32(dbNumber, offset, uint.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Char:
                _memoryStore.Write(dbNumber, offset, [ParseChar(value)]);
                return;
            case PlcValueType.Int:
                _memoryStore.SetInt(dbNumber, offset, short.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Dint:
                _memoryStore.SetDInt(dbNumber, offset, int.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Real:
                _memoryStore.SetReal(dbNumber, offset, float.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.LReal:
                WriteLReal(dbNumber, offset, double.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.S5Time:
                WriteUInt16(dbNumber, offset, EncodeMsToS5Time(ulong.Parse(value.Trim(), CultureInfo.InvariantCulture)));
                return;
            case PlcValueType.Time:
                _memoryStore.SetDInt(dbNumber, offset, int.Parse(value.Trim(), CultureInfo.InvariantCulture));
                return;
            case PlcValueType.Date:
                var date = DateOnly.Parse(value.Trim(), CultureInfo.InvariantCulture);
                WriteUInt16(dbNumber, offset, checked((ushort)(date.DayNumber - S7DateBase.DayNumber)));
                return;
            case PlcValueType.TimeOfDay:
                WriteUInt32(dbNumber, offset, ParseTimeOfDay(value));
                return;
            case PlcValueType.DateAndTime:
                _memoryStore.Write(dbNumber, offset, EncodeDateAndTime(value));
                return;
            case PlcValueType.String:
                SetS7String(dbNumber, offset, stringLength, NormalizeStringStartValue(value));
                return;
            default:
                throw new NotSupportedException($"不支持的数组元素类型: {elementType}");
        }
    }

    private ArrayTypeDefinition? PromptForArrayTypeDefinition(
        PlcValueType defaultElementType,
        int defaultLower,
        int defaultUpper,
        int? defaultStringLength,
        bool allowNegativeIndex)
    {
        var elementOptions = new List<ArrayElementTypeOption>
        {
            new(PlcValueType.Bool, "Bool"),
            new(PlcValueType.Byte, "Byte"),
            new(PlcValueType.Word, "Word"),
            new(PlcValueType.Dword, "DWord"),
            new(PlcValueType.Int, "Int"),
            new(PlcValueType.Dint, "DInt"),
            new(PlcValueType.Real, "Real"),
            new(PlcValueType.LReal, "LReal"),
            new(PlcValueType.String, "String[n]", RequiresLength: true)
        };

        var selected = elementOptions.FirstOrDefault(x => x.Value == defaultElementType)
            ?? elementOptions[0];
        var selectedStringLength = Math.Clamp(defaultStringLength ?? 32, 1, 254);

        var typeCombo = new ComboBox
        {
            MinWidth = 220,
            Margin = new Thickness(0, 8, 0, 10),
            ItemsSource = elementOptions,
            DisplayMemberPath = nameof(ArrayElementTypeOption.Label),
            SelectedItem = selected
        };
        var boundsInput = new TextBox
        {
            Text = $"{defaultLower}..{defaultUpper}",
            MinWidth = 220,
            Margin = new Thickness(0, 8, 0, 4)
        };
        var hintText = new TextBlock
        {
            Text = "示例：0..99 或 0..10",
            Foreground = GetThemeBrush("HintBrush", Brushes.DimGray),
            Margin = new Thickness(0, 0, 0, 8)
        };
        var stringLengthInput = new TextBox
        {
            Text = selectedStringLength.ToString(CultureInfo.InvariantCulture),
            MinWidth = 220,
            Margin = new Thickness(0, 8, 0, 10),
            Visibility = selected.RequiresLength ? Visibility.Visible : Visibility.Collapsed
        };
        var stringLengthLabel = new TextBlock
        {
            Text = "字符串长度 n：",
            Margin = new Thickness(0, 0, 8, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Visibility = selected.RequiresLength ? Visibility.Visible : Visibility.Collapsed
        };
        var previewText = new TextBlock
        {
            Margin = new Thickness(0, 0, 0, 8),
            FontWeight = FontWeights.SemiBold
        };
        var validationText = new TextBlock
        {
            Foreground = GetThemeBrush("DangerBrush", Brushes.Firebrick),
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };

        var okButton = new Button
        {
            Content = "确定",
            Width = 84,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 84,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };

        var form = new Grid();
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var typeLabel = new TextBlock { Text = "数据类型：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var boundsLabel = new TextBlock { Text = "数组限值：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };

        Grid.SetRow(typeLabel, 0);
        Grid.SetColumn(typeLabel, 0);
        Grid.SetRow(typeCombo, 0);
        Grid.SetColumn(typeCombo, 1);

        Grid.SetRow(boundsLabel, 1);
        Grid.SetColumn(boundsLabel, 0);
        Grid.SetRow(boundsInput, 1);
        Grid.SetColumn(boundsInput, 1);

        Grid.SetRow(stringLengthLabel, 2);
        Grid.SetColumn(stringLengthLabel, 0);
        Grid.SetRow(stringLengthInput, 2);
        Grid.SetColumn(stringLengthInput, 1);

        form.Children.Add(typeLabel);
        form.Children.Add(typeCombo);
        form.Children.Add(boundsLabel);
        form.Children.Add(boundsInput);
        form.Children.Add(stringLengthLabel);
        form.Children.Add(stringLengthInput);

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock { Text = "Array[..] of", FontWeight = FontWeights.Bold, FontSize = 14 });
        root.Children.Add(form);
        root.Children.Add(hintText);
        root.Children.Add(previewText);
        root.Children.Add(validationText);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = "Array[..] of",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        ArrayTypeDefinition? result = null;

        void SetValidation(string? message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                validationText.Text = string.Empty;
                validationText.Visibility = Visibility.Collapsed;
                okButton.IsEnabled = true;
                return;
            }

            validationText.Text = message;
            validationText.Visibility = Visibility.Visible;
            okButton.IsEnabled = false;
        }

        bool TryBuildDefinition(out ArrayTypeDefinition definition, out string? error)
        {
            definition = default!;
            error = null;

            if (typeCombo.SelectedItem is not ArrayElementTypeOption option)
            {
                error = "请选择数据类型。";
                return false;
            }

            var boundsRaw = boundsInput.Text.Trim();
            if (!System.Text.RegularExpressions.Regex.IsMatch(boundsRaw, @"^\s*-?\d+\s*\.\.\s*-?\d+\s*$"))
            {
                error = "数组限值格式无效，请输入 lower..upper。";
                return false;
            }

            var parts = boundsRaw.Split("..", StringSplitOptions.TrimEntries);
            if (parts.Length != 2 ||
                !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var lower) ||
                !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var upper))
            {
                error = "数组限值格式无效，请输入 lower..upper。";
                return false;
            }

            if (!allowNegativeIndex && (lower < 0 || upper < 0))
            {
                error = "数组索引不能为负数。";
                return false;
            }

            if (lower > upper)
            {
                error = "下限不能大于上限。";
                return false;
            }

            var count = (long)upper - lower + 1;
            if (count > MaxArrayElementCount)
            {
                error = $"数组元素数量不能超过 {MaxArrayElementCount}。";
                return false;
            }

            int? length = null;
            if (option.RequiresLength)
            {
                if (!int.TryParse(stringLengthInput.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) || n is < 1 or > 254)
                {
                    error = "String[n] 的 n 必须在 1..254。";
                    return false;
                }

                length = n;
            }

            definition = new ArrayTypeDefinition(option.Value, lower, upper, length);
            return true;
        }

        void RefreshPreviewAndValidation()
        {
            if (TryBuildDefinition(out var preview, out var error))
            {
                previewText.Text = $"预览：{BuildArrayTypeEditorText(preview)}";
                SetValidation(null);
            }
            else
            {
                previewText.Text = "预览：Array[?.. ?] of ?";
                SetValidation(error);
            }
        }

        typeCombo.SelectionChanged += (_, _) =>
        {
            var isString = (typeCombo.SelectedItem as ArrayElementTypeOption)?.RequiresLength == true;
            stringLengthLabel.Visibility = isString ? Visibility.Visible : Visibility.Collapsed;
            stringLengthInput.Visibility = isString ? Visibility.Visible : Visibility.Collapsed;
            RefreshPreviewAndValidation();
        };
        boundsInput.TextChanged += (_, _) => RefreshPreviewAndValidation();
        stringLengthInput.TextChanged += (_, _) => RefreshPreviewAndValidation();

        okButton.Click += (_, _) =>
        {
            if (!TryBuildDefinition(out var definition, out var error))
            {
                SetValidation(error);
                return;
            }

            result = definition;
            dialog.DialogResult = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) =>
        {
            dialog.DialogResult = false;
            dialog.Close();
        };

        RefreshPreviewAndValidation();
        typeCombo.Focus();

        return dialog.ShowDialog() == true ? result : null;
    }

    private static string BuildArrayTypeEditorText(ArrayTypeDefinition definition)
    {
        var elementText = definition.ElementType == PlcValueType.String && definition.StringLength.HasValue
            ? $"String[{definition.StringLength.Value}]"
            : definition.ElementType.ToString();
        return $"Array[{definition.LowerBound}..{definition.UpperBound}] of {elementText}";
    }

    private static string NormalizeStringStartValue(string rawValue)
    {
        var value = rawValue.Trim();
        return value == "''" ? string.Empty : rawValue;
    }

    private void RecalculateOffsetsForDb(int dbNumber)
    {
        var rows = Variables.Where(x => x.DbNumber == dbNumber).ToList();
        var byteOffset = 0;
        var bitOffset = 0;

        foreach (var row in rows)
        {
            if (row.Type == PlcValueType.Bool)
            {
                row.Offset = byteOffset;
                row.Bit = bitOffset;

                bitOffset++;
                if (bitOffset >= 8)
                {
                    bitOffset = 0;
                    byteOffset++;
                }

                continue;
            }

            if (bitOffset > 0)
            {
                byteOffset++;
                bitOffset = 0;
            }

            var (size, alignment) = GetTypeLayout(row);
            byteOffset = AlignOffset(byteOffset, alignment);
            row.Offset = byteOffset;
            row.Bit = 0;
            byteOffset += size;
        }

        EnsureVariablesSortOrder();
        RefreshFilteredVariablesSafely();
    }

    private void EnsureVariablesSortOrder()
    {
        if (FilteredVariables is IEditableCollectionView editable &&
            (editable.IsAddingNew || editable.IsEditingItem))
        {
            return;
        }

        var sorts = FilteredVariables.SortDescriptions;
        if (sorts.Count == 2 &&
            sorts[0].PropertyName == nameof(VariableRow.Offset) &&
            sorts[0].Direction == ListSortDirection.Ascending &&
            sorts[1].PropertyName == nameof(VariableRow.Bit) &&
            sorts[1].Direction == ListSortDirection.Ascending)
        {
            return;
        }

        sorts.Clear();
        sorts.Add(new SortDescription(nameof(VariableRow.Offset), ListSortDirection.Ascending));
        sorts.Add(new SortDescription(nameof(VariableRow.Bit), ListSortDirection.Ascending));
    }

    private void RefreshFilteredVariablesSafely()
    {
        if (FilteredVariables is IEditableCollectionView editable &&
            (editable.IsAddingNew || editable.IsEditingItem))
        {
            return;
        }

        FilteredVariables.Refresh();
    }

    private (int Size, int Alignment) GetTypeLayout(VariableRow row)
    {
        if (row.Type == PlcValueType.Array)
        {
            var (elementSize, elementAlignment) = GetScalarTypeLayout(row.ArrayElementType, row.StringLength);
            return (elementSize * row.ArrayLength, elementAlignment);
        }

        return row.Type switch
        {
            PlcValueType.Byte => (1, 1),
            PlcValueType.Usint => (1, 1),
            PlcValueType.Sint => (1, 1),
            PlcValueType.Char => (1, 1),
            PlcValueType.Word => (2, 2),
            PlcValueType.Uint => (2, 2),
            PlcValueType.Int => (2, 2),
            PlcValueType.Date => (2, 2),
            PlcValueType.S5Time => (2, 2),
            PlcValueType.Dword => (4, 4),
            PlcValueType.Udint => (4, 4),
            PlcValueType.Dint => (4, 4),
            PlcValueType.Real => (4, 4),
            PlcValueType.LReal => (8, 8),
            PlcValueType.Time => (4, 4),
            PlcValueType.TimeOfDay => (4, 4),
            PlcValueType.DateAndTime => (8, 8),
            PlcValueType.String => (Math.Clamp(row.StringLength, 1, 254) + 2, 2),
            PlcValueType.WString => (4 + Math.Clamp(row.StringLength, 1, 1024) * 2, 2),
            _ => (1, 1)
        };
    }

    private (int Size, int Alignment) GetScalarTypeLayout(PlcValueType type, int stringLength)
    {
        return type switch
        {
            PlcValueType.Bool => (1, 1),
            PlcValueType.Byte => (1, 1),
            PlcValueType.Usint => (1, 1),
            PlcValueType.Sint => (1, 1),
            PlcValueType.Char => (1, 1),
            PlcValueType.Word => (2, 2),
            PlcValueType.Uint => (2, 2),
            PlcValueType.Int => (2, 2),
            PlcValueType.Date => (2, 2),
            PlcValueType.S5Time => (2, 2),
            PlcValueType.Dword => (4, 4),
            PlcValueType.Udint => (4, 4),
            PlcValueType.Dint => (4, 4),
            PlcValueType.Real => (4, 4),
            PlcValueType.LReal => (8, 8),
            PlcValueType.Time => (4, 4),
            PlcValueType.TimeOfDay => (4, 4),
            PlcValueType.DateAndTime => (8, 8),
            PlcValueType.String => (Math.Clamp(stringLength, 1, 254) + 2, 2),
            PlcValueType.WString => (4 + Math.Clamp(stringLength, 1, 1024) * 2, 2),
            _ => throw new NotSupportedException($"不支持的数组元素类型: {type}")
        };
    }

    private static int AlignOffset(int value, int alignment)
    {
        if (alignment <= 1)
        {
            return value;
        }

        var mod = value % alignment;
        return mod == 0 ? value : value + (alignment - mod);
    }

    private ushort ReadUInt16(int db, int offset)
    {
        var bytes = _memoryStore.Read(db, offset, 2);
        return BinaryPrimitives.ReadUInt16BigEndian(bytes);
    }

    private uint ReadUInt32(int db, int offset)
    {
        var bytes = _memoryStore.Read(db, offset, 4);
        return BinaryPrimitives.ReadUInt32BigEndian(bytes);
    }

    private void WriteUInt16(int db, int offset, ushort value)
    {
        Span<byte> bytes = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, value);
        _memoryStore.Write(db, offset, bytes.ToArray());
    }

    private void WriteUInt32(int db, int offset, uint value)
    {
        Span<byte> bytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(bytes, value);
        _memoryStore.Write(db, offset, bytes.ToArray());
    }

    private double ReadLReal(int db, int offset)
    {
        var bytes = _memoryStore.Read(db, offset, 8);
        if (BitConverter.IsLittleEndian)
        {
            Array.Reverse(bytes);
        }

        return BitConverter.ToDouble(bytes, 0);
    }

    private void WriteLReal(int db, int offset, double value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
        {
            Array.Reverse(bytes);
        }

        _memoryStore.Write(db, offset, bytes);
    }

    private string GetS7String(int db, int offset)
    {
        var header = _memoryStore.Read(db, offset, 2);
        var maxLength = header[0];
        var currentLength = Math.Min(header[1], maxLength);
        if (currentLength == 0)
        {
            return string.Empty;
        }

        var payload = _memoryStore.Read(db, offset + 2, currentLength);
        return Encoding.ASCII.GetString(payload);
    }

    private void SetS7String(int db, int offset, int length, string value)
    {
        var maxLength = Math.Clamp(length, 1, 254);
        var text = NormalizeStringStartValue(value);
        var payload = Encoding.ASCII.GetBytes(text);
        if (payload.Length > maxLength)
        {
            payload = payload[..maxLength];
        }

        var raw = new byte[maxLength + 2];
        raw[0] = (byte)maxLength;
        raw[1] = (byte)payload.Length;
        Buffer.BlockCopy(payload, 0, raw, 2, payload.Length);
        _memoryStore.Write(db, offset, raw);
    }

    private string GetWString(int db, int offset)
    {
        var header = _memoryStore.Read(db, offset, 4);
        var maxLength = BinaryPrimitives.ReadUInt16BigEndian(header.AsSpan(0, 2));
        var currentLength = BinaryPrimitives.ReadUInt16BigEndian(header.AsSpan(2, 2));
        var charLength = (int)Math.Min(currentLength, maxLength);
        if (charLength == 0)
        {
            return string.Empty;
        }

        var payload = _memoryStore.Read(db, offset + 4, charLength * 2);
        return Encoding.BigEndianUnicode.GetString(payload);
    }

    private void SetWString(int db, int offset, int length, string value)
    {
        var maxLength = Math.Clamp(length, 1, 1024);
        var text = NormalizeStringStartValue(value);
        if (text.Length > maxLength)
        {
            text = text[..maxLength];
        }

        var payload = Encoding.BigEndianUnicode.GetBytes(text);
        var raw = new byte[4 + maxLength * 2];
        BinaryPrimitives.WriteUInt16BigEndian(raw.AsSpan(0, 2), (ushort)maxLength);
        BinaryPrimitives.WriteUInt16BigEndian(raw.AsSpan(2, 2), (ushort)text.Length);
        Buffer.BlockCopy(payload, 0, raw, 4, payload.Length);
        _memoryStore.Write(db, offset, raw);
    }

    private static byte ParseChar(string raw)
    {
        if (string.IsNullOrEmpty(raw))
        {
            return 0;
        }

        return Encoding.ASCII.GetBytes([raw[0]])[0];
    }

    private static string FormatTimeOfDay(uint milliseconds)
    {
        var clamped = milliseconds % 86_400_000u;
        var ts = TimeSpan.FromMilliseconds(clamped);
        return $"{(int)ts.TotalHours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds:000}";
    }

    private static uint ParseTimeOfDay(string raw)
    {
        var value = raw.Trim();
        if (uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ms))
        {
            return ms % 86_400_000u;
        }

        var ts = TimeSpan.Parse(value, CultureInfo.InvariantCulture);
        var totalMs = ts.TotalMilliseconds;
        if (totalMs < 0)
        {
            throw new FormatException("TIME_OF_DAY 涓嶈兘涓鸿礋鏁般€");
        }

        return (uint)((ulong)totalMs % 86_400_000ul);
    }

    private static ushort EncodeMsToS5Time(ulong totalMs)
    {
        var factors = new[] { 10ul, 100ul, 1_000ul, 10_000ul };
        for (var baseCode = 0; baseCode < factors.Length; baseCode++)
        {
            var value = totalMs / factors[baseCode];
            if (value <= 999)
            {
                var bcd = (int)value;
                var hundreds = bcd / 100;
                var tens = (bcd / 10) % 10;
                var ones = bcd % 10;
                return (ushort)((baseCode << 12) | (hundreds << 8) | (tens << 4) | ones);
            }
        }

        throw new FormatException("S5TIME 瓒呭嚭鍙紪鐮佽寖鍥淬€");
    }

    private static int DecodeS5TimeToMs(ushort raw)
    {
        var baseCode = (raw >> 12) & 0xF;
        var hundreds = (raw >> 8) & 0xF;
        var tens = (raw >> 4) & 0xF;
        var ones = raw & 0xF;
        var value = hundreds * 100 + tens * 10 + ones;
        var factor = baseCode switch
        {
            0 => 10,
            1 => 100,
            2 => 1_000,
            3 => 10_000,
            _ => throw new FormatException("S5TIME 基数无效。")
        };
        return value * factor;
    }

    private static string DecodeDateAndTime(byte[] bytes)
    {
        if (bytes.Length != 8)
        {
            throw new FormatException("DATE_AND_TIME 闀垮害蹇呴』鏄?8 瀛楄妭銆");
        }

        var year2 = FromBcd(bytes[0]);
        var year = year2 >= 90 ? 1900 + year2 : 2000 + year2;
        var month = FromBcd(bytes[1]);
        var day = FromBcd(bytes[2]);
        var hour = FromBcd(bytes[3]);
        var minute = FromBcd(bytes[4]);
        var second = FromBcd(bytes[5]);
        var millisecond = FromBcd(bytes[6]) * 10 + ((bytes[7] >> 4) & 0x0F);

        var dt = new DateTime(year, month, day, hour, minute, second, millisecond);
        return dt.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
    }

    private static byte[] EncodeDateAndTime(string raw)
    {
        var value = raw.Trim();
        DateTime dt;
        if (!DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) &&
            !DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
        {
            dt = DateTime.Parse(value, CultureInfo.InvariantCulture);
        }

        var year2 = dt.Year % 100;
        var dow = (int)dt.DayOfWeek;
        if (dow == 0)
        {
            dow = 7;
        }

        var msHigh = dt.Millisecond / 10;
        var msLow = dt.Millisecond % 10;

        return
        [
            ToBcd(year2),
            ToBcd(dt.Month),
            ToBcd(dt.Day),
            ToBcd(dt.Hour),
            ToBcd(dt.Minute),
            ToBcd(dt.Second),
            ToBcd(msHigh),
            (byte)((msLow << 4) | (dow & 0x0F))
        ];
    }

    private static byte ToBcd(int value)
    {
        if (value is < 0 or > 99)
        {
            throw new FormatException("BCD 鏁板€艰寖鍥村簲涓?0-99銆");
        }

        return (byte)(((value / 10) << 4) | (value % 10));
    }

    private static int FromBcd(byte value)
    {
        return ((value >> 4) & 0x0F) * 10 + (value & 0x0F);
    }

    private static Dictionary<int, int>? LoadDbConfig(string path)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        var text = File.ReadAllText(path);
        var raw = JsonSerializer.Deserialize<Dictionary<string, int>>(text);
        if (raw is null || raw.Count == 0)
        {
            return null;
        }

        var result = new Dictionary<int, int>();
        foreach (var item in raw)
        {
            var key = item.Key.Trim().ToUpperInvariant();
            if (!key.StartsWith("DB") || !int.TryParse(key[2..], out var dbNumber))
            {
                continue;
            }

            result[dbNumber] = item.Value;
        }

        return result.Count == 0 ? null : result;
    }

    private static SimulatorState? LoadAppState(string path)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            var text = File.ReadAllText(path);
            return JsonSerializer.Deserialize<SimulatorState>(text);
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<int, int> BuildDbConfigFromState(SimulatorState state)
    {
        var result = new Dictionary<int, int>();
        if (state.DbBlocks is not null)
        {
            foreach (var item in state.DbBlocks)
            {
                if (item.Key > 0)
                {
                    result[item.Key] = Math.Max(item.Value?.Length ?? 0, 1);
                }
            }
        }

        return result.Count > 0 ? result : new Dictionary<int, int> { [1] = DefaultDbSize };
    }

    private bool TryRestoreAppState(SimulatorState state)
    {
        if (state.ProjectNodes is null || state.ProjectNodes.Count == 0)
        {
            return false;
        }

        ProjectNodes.Clear();
        Variables.Clear();
        _plcEndpointSettings.Clear();

        foreach (var nodeState in state.ProjectNodes)
        {
            var node = BuildProjectNodeFromState(nodeState, null);
            if (node is not null)
            {
                ProjectNodes.Add(node);
            }
        }
        SortPlcRootsByIndex();

        if (ProjectNodes.Count == 0)
        {
            return false;
        }

        foreach (var root in ProjectNodes)
        {
            if (root.NodeType != ProjectNodeType.PlcRoot)
            {
                continue;
            }

            if (state.PlcSettings is not null &&
                state.PlcSettings.TryGetValue(root.Name, out var setting) &&
                !string.IsNullOrWhiteSpace(setting.Address))
            {
                _plcEndpointSettings[root] = new PlcEndpointSettings(setting.Address, setting.Port, setting.AutoStart);
            }
            else
            {
                var defaultAddress = _plcEndpointSettings.Count == 0 ? GetPreferredDefaultAddress() : string.Empty;
                _plcEndpointSettings[root] = new PlcEndpointSettings(defaultAddress, DefaultPlcPort, true);
            }
        }
        if (state.DbBlocks is not null)
        {
            foreach (var item in state.DbBlocks)
            {
                if (item.Key <= 0 || item.Value is null)
                {
                    continue;
                }

                var target = _memoryStore.EnsureDb(item.Key, Math.Max(item.Value.Length, 1), out _);
                Buffer.BlockCopy(item.Value, 0, target, 0, Math.Min(item.Value.Length, target.Length));
                foreach (var host in _plcServerHosts.Values)
                {
                    host.EnsureDbAreaRegistered(item.Key);
                }
            }
        }

        if (state.Variables is not null)
        {
            foreach (var rowState in state.Variables)
            {
                if (rowState.DbNumber <= 0)
                {
                    continue;
                }

                var row = new VariableRow
                {
                    Name = rowState.Name,
                    DbNumber = rowState.DbNumber,
                    Offset = rowState.Offset,
                    Bit = rowState.Bit,
                    Type = rowState.Type,
                    StringLength = rowState.StringLength,
                    ArrayElementType = rowState.ArrayElementType,
                    ArrayLowerBound = rowState.ArrayLowerBound,
                    ArrayUpperBound = rowState.ArrayUpperBound,
                    StartValue = rowState.StartValue,
                    Value = rowState.Value
                };
                Variables.Add(row);
                EnsureDbMemoryReady(row.DbNumber);
            }
        }

        _selectedNode = FindFirstDataBlockNode() ?? ProjectNodes.FirstOrDefault();
        return true;
    }

    private void SaveAppState(string path)
    {
        try
        {
            var state = new SimulatorState
            {
                ProjectNodes = ProjectNodes.Select(BuildProjectNodeState).ToList(),
                Variables = Variables.Select(row => new VariableRowState
                {
                    Name = row.Name,
                    DbNumber = row.DbNumber,
                    Offset = row.Offset,
                    Bit = row.Bit,
                    Type = row.Type,
                    StringLength = row.StringLength,
                    ArrayElementType = row.ArrayElementType,
                    ArrayLowerBound = row.ArrayLowerBound,
                    ArrayUpperBound = row.ArrayUpperBound,
                    StartValue = row.StartValue,
                    Value = row.Value
                }).ToList(),
                DbBlocks = _memoryStore.GetDbBlocksSnapshot().ToDictionary(
                    x => x.Key,
                    x => x.Value.ToArray()),
                PlcSettings = _plcEndpointSettings
                    .Where(x => x.Key.NodeType == ProjectNodeType.PlcRoot)
                    .ToDictionary(
                        x => x.Key.Name,
                        x => new PlcEndpointSettingState
                        {
                            Address = x.Value.Address,
                            Port = x.Value.Port,
                            AutoStart = x.Value.AutoStart
                        })
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            File.WriteAllText(path, JsonSerializer.Serialize(state, options));
        }
        catch (Exception ex)
        {
            AddLog($"保存状态失败: {ex.Message}");
        }
    }

    private static ProjectNodeState BuildProjectNodeState(ProjectTreeNode node)
    {
        return new ProjectNodeState
        {
            Name = node.Name,
            NodeType = node.NodeType,
            Children = node.Children.Select(BuildProjectNodeState).ToList()
        };
    }

    private static ProjectTreeNode? BuildProjectNodeFromState(ProjectNodeState state, ProjectTreeNode? parent)
    {
        if (string.IsNullOrWhiteSpace(state.Name))
        {
            return null;
        }

        var node = new ProjectTreeNode
        {
            Name = state.Name,
            NodeType = state.NodeType,
            Parent = parent
        };

        if (state.Children is not null)
        {
            foreach (var childState in state.Children)
            {
                var child = BuildProjectNodeFromState(childState, node);
                if (child is not null)
                {
                    node.Children.Add(child);
                }
            }
        }

        return node;
    }

    private ProjectTreeNode? FindFirstDataBlockNode()
    {
        foreach (var root in ProjectNodes)
        {
            var dataBlock = FindNodeByType(root, ProjectNodeType.DataBlock);
            if (dataBlock is not null)
            {
                return dataBlock;
            }
        }

        return null;
    }

    private MenuItem CreateMenuItem(string header, RoutedEventHandler onClick)
    {
        var item = new MenuItem { Header = header };
        if (TryFindResource("ContextMenuItemStyle") is Style menuItemStyle)
        {
            item.Style = menuItemStyle;
        }

        item.Click += onClick;
        return item;
    }

    private void ApplyContextMenuStyle(ContextMenu menu)
    {
        if (TryFindResource(typeof(ContextMenu)) is Style contextMenuStyle)
        {
            menu.Style = contextMenuStyle;
        }
    }

    private Separator CreateContextMenuSeparator()
    {
        var separator = new Separator();
        if (TryFindResource("ContextMenuSeparatorStyle") is Style separatorStyle)
        {
            separator.Style = separatorStyle;
        }

        return separator;
    }

    private PlcEndpointSettings GetOrCreatePlcSettings(ProjectTreeNode plcRoot)
    {
        if (_plcEndpointSettings.TryGetValue(plcRoot, out var settings))
        {
            return settings;
        }

        settings = new PlcEndpointSettings(GetPreferredDefaultAddress(), DefaultPlcPort, true);
        _plcEndpointSettings[plcRoot] = settings;
        return settings;
    }

    private void ApplyPlcSettings(ProjectTreeNode plcRoot, PlcEndpointSettings settings)
    {
        if (settings.AutoStart)
        {
            if (!TryValidateIpAddress(settings.Address, out var ipError))
            {
                AddLog($"[{plcRoot.Name}] 自动启动跳过: {ipError}");
                StopPlcServer(plcRoot);
                UpdateConnectionStatus(_plcServerHosts.Count > 0);
                return;
            }

            StartPlcServer(plcRoot, settings);
        }
        else
        {
            StopPlcServer(plcRoot);
        }

        UpdateConnectionStatus(_plcServerHosts.Count > 0);
    }

    private void StartPlcServer(ProjectTreeNode plcRoot, PlcEndpointSettings settings)
    {
        StopPlcServer(plcRoot);
        try
        {
            if (!TryResolveBindAddress(settings.Address, out var configuredIp, out var bindIp, out var bindError))
            {
                throw new InvalidOperationException(bindError);
            }

            var bindCandidates = new List<IPAddress> { bindIp };
            if (!IPAddress.Any.Equals(bindIp))
            {
                bindCandidates.Add(IPAddress.Any);
            }

            S7ServerHost? startedHost = null;
            IPAddress? actualBindIp = null;
            Exception? lastStartException = null;
            foreach (var candidate in bindCandidates)
            {
                try
                {
                    var host = new S7ServerHost(_memoryStore, AddLog);
                    host.Start(candidate.ToString(), settings.Port);
                    startedHost = host;
                    actualBindIp = candidate;
                    break;
                }
                catch (Exception ex)
                {
                    lastStartException = ex;
                }
            }

            if (startedHost is null || actualBindIp is null)
            {
                throw lastStartException ?? new InvalidOperationException("未知启动错误。");
            }

            _plcServerHosts[plcRoot] = startedHost;
            if (configuredIp.Equals(actualBindIp))
            {
                AddLog($"[{plcRoot.Name}] 已启动 {configuredIp}:{settings.Port}");
            }
            else if (IPAddress.Any.Equals(actualBindIp))
            {
                AddLog($"[{plcRoot.Name}] 已启动 配置IP={configuredIp}:{settings.Port}，绑定地址=0.0.0.0（自动回退）");
            }
            else
            {
                AddLog($"[{plcRoot.Name}] 已启动 配置IP={configuredIp}:{settings.Port}，绑定网卡={actualBindIp}");
            }

            UpdateSimulatorRunState();
        }
        catch (Exception ex)
        {
            var message = BuildFriendlyPlcStartError(settings.Address, settings.Port, ex);
            AddLog($"[{plcRoot.Name}] 启动失败: {message}");
            MessageBox.Show(this, message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            UpdateSimulatorRunState();
        }
    }

    private void StopPlcServer(ProjectTreeNode plcRoot)
    {
        if (!_plcServerHosts.TryGetValue(plcRoot, out var host))
        {
            return;
        }

        host.Stop();
        _plcServerHosts.Remove(plcRoot);
        AddLog($"[{plcRoot.Name}] 已停止");
        UpdateSimulatorRunState();
    }

    private PlcEndpointSettings? PromptForPlcSettings(ProjectTreeNode plcRoot, PlcEndpointSettings current)
    {
        var plcName = plcRoot.Name;
        var nicHintToolTip = new ToolTip
        {
            Content = $"识别到的本机网卡:\n{BuildLocalNicHintMultiline()}",
            MaxWidth = 420
        };
        var ipInput = new TextBox
        {
            Text = current.Address,
            MinWidth = 220,
            Margin = new Thickness(0, 8, 0, 10),
            ToolTip = nicHintToolTip
        };
        ToolTipService.SetInitialShowDelay(ipInput, 0);
        ToolTipService.SetShowDuration(ipInput, 60000);
        ToolTipService.SetPlacement(ipInput, System.Windows.Controls.Primitives.PlacementMode.Right);
        var portInput = new TextBox
        {
            Text = current.Port.ToString(CultureInfo.InvariantCulture),
            MinWidth = 220,
            Margin = new Thickness(0, 0, 0, 10)
        };
        var autoStartCheckBox = new CheckBox
        {
            Content = "自动启动",
            IsChecked = current.AutoStart,
            Margin = new Thickness(0, 0, 0, 12)
        };

        var form = new Grid();
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var ipLabel = new TextBlock { Text = "IP 地址：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var portLabel = new TextBlock { Text = "端口：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };

        Grid.SetRow(ipLabel, 0);
        Grid.SetColumn(ipLabel, 0);
        Grid.SetRow(ipInput, 0);
        Grid.SetColumn(ipInput, 1);
        Grid.SetRow(portLabel, 1);
        Grid.SetColumn(portLabel, 0);
        Grid.SetRow(portInput, 1);
        Grid.SetColumn(portInput, 1);
        Grid.SetRow(autoStartCheckBox, 2);
        Grid.SetColumn(autoStartCheckBox, 1);

        form.Children.Add(ipLabel);
        form.Children.Add(ipInput);
        form.Children.Add(portLabel);
        form.Children.Add(portInput);
        form.Children.Add(autoStartCheckBox);

        var okButton = new Button
        {
            Content = "确定",
            Width = 84,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 84,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };
        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(form);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = $"{plcName} 设置",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        PlcEndpointSettings? updated = null;
        okButton.Click += (_, _) =>
        {
            var address = ipInput.Text.Trim();
            if (!TryResolveBindAddress(address, out _, out _, out var ipError))
            {
                var message = $"输入的 IP 不在识别网段内，无法确认。\n{ipError}\n\n识别到的本机网卡:\n{BuildLocalNicHintMultiline()}";
                MessageBox.Show(dialog, message, "IP 校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (TryFindIpConflict(plcRoot, address, out var conflictPlcName))
            {
                MessageBox.Show(dialog, $"IP 已被占用：{address}\n已被模拟器 [{conflictPlcName}] 使用。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(portInput.Text.Trim(), out var port) || port is < 1 or > 65535)
            {
                MessageBox.Show(dialog, "端口范围必须在 1-65535。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            updated = new PlcEndpointSettings(address, port, autoStartCheckBox.IsChecked == true);
            dialog.DialogResult = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) =>
        {
            dialog.DialogResult = false;
            dialog.Close();
        };

        var result = dialog.ShowDialog();
        return result == true ? updated : null;
    }

    private static bool TryValidateIpAddress(string address, out string error)
    {
        return TryResolveBindAddress(address, out _, out _, out error);
    }

    private bool TryFindIpConflict(ProjectTreeNode currentPlcRoot, string address, out string conflictPlcName)
    {
        conflictPlcName = string.Empty;
        if (!IPAddress.TryParse(address, out var targetIp))
        {
            return false;
        }

        foreach (var item in _plcEndpointSettings)
        {
            if (ReferenceEquals(item.Key, currentPlcRoot))
            {
                continue;
            }

            if (!IPAddress.TryParse(item.Value.Address, out var existingIp))
            {
                continue;
            }

            if (!targetIp.Equals(existingIp))
            {
                continue;
            }

            conflictPlcName = item.Key.Name;
            return true;
        }

        return false;
    }

    private static bool TryResolveBindAddress(string configuredAddress, out IPAddress configuredIp, out IPAddress bindIp, out string error)
    {
        configuredIp = IPAddress.None;
        bindIp = IPAddress.None;

        if (string.IsNullOrWhiteSpace(configuredAddress))
        {
            error = "IP 地址不能为空。";
            return false;
        }

        if (!IPAddress.TryParse(configuredAddress.Trim(), out var parsedIp) || parsedIp is null)
        {
            error = "IP 地址格式无效，请输入合法 IP（例如 192.168.1.100）。";
            return false;
        }
        configuredIp = parsedIp;

        if (configuredIp.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
        {
            error = "当前仅支持 IPv4 地址。";
            return false;
        }

        if (IPAddress.Any.Equals(configuredIp) || IPAddress.Loopback.Equals(configuredIp) || IsLoopbackV4(configuredIp))
        {
            bindIp = configuredIp;
            error = string.Empty;
            return true;
        }

        var interfaces = GetLocalIpv4Interfaces();

        // Prefer routing result, but only if it belongs to supported local NICs.
        if (TryGetRouteLocalIpv4(configuredIp, out var routedLocalIp) &&
            interfaces.Any(x => x.Address.Equals(routedLocalIp)))
        {
            bindIp = routedLocalIp;
            error = string.Empty;
            return true;
        }

        foreach (var local in interfaces)
        {
            if (local.Address.Equals(configuredIp))
            {
                bindIp = local.Address;
                error = string.Empty;
                return true;
            }
        }

        foreach (var local in interfaces)
        {
            if (IsSameSubnet(configuredIp, local.Address, local.Mask))
            {
                bindIp = local.Address;
                error = string.Empty;
                return true;
            }
        }

        foreach (var local in interfaces)
        {
            if (IsLooseSamePrivateNetwork(configuredIp, local.Address))
            {
                bindIp = local.Address;
                error = string.Empty;
                return true;
            }
        }

        error = $"IP 地址 {configuredIp} 未与本机任一可用网卡处于同一网段，无法启动。\n本机网卡地址: {BuildLocalAddressHint()}";
        return false;
    }

    private static bool IsPortOccupiedForBinding(IPAddress bindIp, int port, out string hint)
    {
        var listeners = IPGlobalProperties.GetIPGlobalProperties().GetActiveTcpListeners();
        var occupied = listeners.Where(x => x.Port == port).ToList();
        if (occupied.Count == 0)
        {
            hint = string.Empty;
            return false;
        }

        if (IPAddress.Any.Equals(bindIp))
        {
            hint = $"0.0.0.0:{port} 监听会与现有监听冲突（{string.Join(", ", occupied.Select(x => $"{x.Address}:{x.Port}"))}）。";
            return true;
        }

        var conflict = occupied.Any(x => IPAddress.Any.Equals(x.Address) || x.Address.Equals(bindIp));
        if (!conflict)
        {
            hint = string.Empty;
            return false;
        }

        hint = $"{bindIp}:{port} 或 0.0.0.0:{port} 已存在监听（{string.Join(", ", occupied.Select(x => $"{x.Address}:{x.Port}"))}）。";
        return true;
    }

    private static bool TryGetRouteLocalIpv4(IPAddress target, out IPAddress localIp)
    {
        localIp = IPAddress.None;
        if (target.AddressFamily != AddressFamily.InterNetwork)
        {
            return false;
        }

        try
        {
            using var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            socket.Connect(new IPEndPoint(target, 65530));
            if (socket.LocalEndPoint is not IPEndPoint endpoint)
            {
                return false;
            }

            if (IPAddress.Any.Equals(endpoint.Address) || IPAddress.Loopback.Equals(endpoint.Address))
            {
                return false;
            }

            localIp = endpoint.Address;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static List<(IPAddress Address, IPAddress Mask)> GetLocalIpv4Interfaces()
    {
        return GetLocalIpv4InterfaceDetails()
            .Select(x => (x.Address, x.Mask))
            .ToList();
    }

    private static List<(string Name, IPAddress Address, IPAddress Mask)> GetLocalIpv4InterfaceDetails()
    {
        var result = new List<(string Name, IPAddress Address, IPAddress Mask)>();
        foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (!IsSupportedLocalNic(nic))
            {
                continue;
            }

            foreach (var uni in nic.GetIPProperties().UnicastAddresses)
            {
                if (uni.Address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    continue;
                }

                var mask = uni.IPv4Mask;
                if (mask is null && uni.PrefixLength is > 0 and <= 32)
                {
                    mask = BuildIpv4MaskFromPrefixLength(uni.PrefixLength);
                }

                if (mask is null)
                {
                    continue;
                }

                result.Add((nic.Name, uni.Address, mask));
            }
        }

        return result;
    }

    private static bool IsSupportedLocalNic(NetworkInterface nic)
    {
        if (nic.OperationalStatus != OperationalStatus.Up)
        {
            return false;
        }

        if (nic.NetworkInterfaceType is NetworkInterfaceType.Loopback or NetworkInterfaceType.Tunnel)
        {
            return false;
        }

        var text = $"{nic.Name} {nic.Description}".ToLowerInvariant();
        var virtualKeywords = new[]
        {
            "virtual", "vmware", "virtualbox", "vbox", "hyper-v", "vethernet",
            "host-only", "loopback", "tap", "tun", "wireguard", "tailscale",
            "zerotier", "npcap", "wintun"
        };

        return virtualKeywords.All(keyword => !text.Contains(keyword, StringComparison.Ordinal));
    }

    private static string BuildLocalNicHintMultiline()
    {
        var lines = GetAllIpv4InterfaceDetails()
            .Select(x => $"{x.Name}:{x.Address}")
            .Distinct()
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (lines.Count == 0)
        {
            return "未识别到可用网卡";
        }

        return string.Join("\n", lines);
    }

    private static List<(string Name, IPAddress Address)> GetAllIpv4InterfaceDetails()
    {
        var result = new List<(string Name, IPAddress Address)>();
        foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
        {
            foreach (var uni in nic.GetIPProperties().UnicastAddresses)
            {
                if (uni.Address.AddressFamily != AddressFamily.InterNetwork)
                {
                    continue;
                }

                if (IsLoopbackV4(uni.Address))
                {
                    continue;
                }

                result.Add((nic.Name, uni.Address));
            }
        }

        return result;
    }

    private static bool IsLoopbackV4(IPAddress ip)
    {
        if (IPAddress.Loopback.Equals(ip))
        {
            return true;
        }

        var bytes = ip.GetAddressBytes();
        return bytes.Length == 4 && bytes[0] == 127;
    }

    private static string GetPreferredDefaultAddress()
    {
        var first = GetLocalIpv4Interfaces()
            .Select(x => x.Address)
            .FirstOrDefault();
        return first?.ToString() ?? string.Empty;
    }

    private static IPAddress BuildIpv4MaskFromPrefixLength(int prefixLength)
    {
        var bytes = new byte[4];
        var remain = prefixLength;
        for (var i = 0; i < bytes.Length; i++)
        {
            if (remain >= 8)
            {
                bytes[i] = 0xFF;
                remain -= 8;
                continue;
            }

            if (remain > 0)
            {
                bytes[i] = (byte)(0xFF << (8 - remain));
                remain = 0;
            }
        }

        return new IPAddress(bytes);
    }

    private static string BuildLocalAddressHint()
    {
        var addresses = GetLocalIpv4Interfaces()
            .Select(x => x.Address.ToString())
            .Distinct()
            .OrderBy(x => x)
            .ToList();
        if (addresses.Count == 0)
        {
            return "0.0.0.0";
        }

        var preview = string.Join(", ", addresses.Take(6));
        return addresses.Count > 6 ? $"{preview} ..." : preview;
    }

    private static bool IsSameSubnet(IPAddress target, IPAddress local, IPAddress mask)
    {
        var targetBytes = target.GetAddressBytes();
        var localBytes = local.GetAddressBytes();
        var maskBytes = mask.GetAddressBytes();
        if (targetBytes.Length != 4 || localBytes.Length != 4 || maskBytes.Length != 4)
        {
            return false;
        }

        for (var i = 0; i < 4; i++)
        {
            if ((targetBytes[i] & maskBytes[i]) != (localBytes[i] & maskBytes[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsLooseSamePrivateNetwork(IPAddress a, IPAddress b)
    {
        var ab = a.GetAddressBytes();
        var bb = b.GetAddressBytes();
        if (ab.Length != 4 || bb.Length != 4)
        {
            return false;
        }

        if (ab[0] == 10 && bb[0] == 10)
        {
            return ab[1] == bb[1];
        }

        if (ab[0] == 172 && bb[0] == 172 && ab[1] is >= 16 and <= 31 && bb[1] is >= 16 and <= 31)
        {
            return ab[1] == bb[1];
        }

        if (ab[0] == 192 && ab[1] == 168 && bb[0] == 192 && bb[1] == 168)
        {
            return true;
        }

        return false;
    }

    private static string BuildFriendlyPlcStartError(string address, int port, Exception ex)
    {
        var raw = ex.Message;
        if (raw.Contains("Can't assign requested address", StringComparison.OrdinalIgnoreCase))
        {
            return $"PLC 启动失败：无法绑定 {address}:{port}。\n原因：配置 IP 既不是本机地址，也不在可用网卡同网段。\n建议：使用与本机网卡同网段的 IP。\n本机网卡地址: {BuildLocalAddressHint()}\n\n原始错误: {raw}";
        }

        if (raw.Contains("Address already in use", StringComparison.OrdinalIgnoreCase))
        {
            return $"PLC 启动失败：{address}:{port} 已被占用。\n建议：更换端口或停止占用该端口的进程。\n\n原始错误: {raw}";
        }

        if (raw.Contains("access permissions", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("Permission denied", StringComparison.OrdinalIgnoreCase))
        {
            return $"PLC 启动失败：权限不足，无法监听 {address}:{port}。\n建议：使用管理员权限运行，或改用高位端口（例如 1102）。\n\n原始错误: {raw}";
        }

        return $"PLC 启动失败：{raw}";
    }

    private string? PromptForName(string currentName)
    {
        var input = new TextBox
        {
            Text = currentName,
            MinWidth = 260,
            Margin = new Thickness(0, 10, 0, 12)
        };

        var okButton = new Button
        {
            Content = "纭畾",
            Width = 80,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 80,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttons.Children.Add(okButton);
        buttons.Children.Add(cancelButton);

        var root = new StackPanel
        {
            Margin = new Thickness(16)
        };
        root.Children.Add(new TextBlock { Text = "请输入新的节点名称：" });
        root.Children.Add(input);
        root.Children.Add(buttons);

        var dialog = new Window
        {
            Title = "重命名节点",
            Content = root,
            ResizeMode = ResizeMode.NoResize,
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this
        };
        ApplyHtmlDialogTheme(dialog);

        okButton.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(input.Text))
            {
                MessageBox.Show(dialog, "名称不能为空。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            dialog.DialogResult = true;
            dialog.Close();
        };

        input.Focus();
        input.SelectAll();

        var result = dialog.ShowDialog();
        return result == true ? input.Text : null;
    }

    private string? PromptForDbBlockRename(string currentDisplayName)
    {
        if (!TryParseDbDisplayName(currentDisplayName, out var currentCustomName, out var dbNumber))
        {
            return PromptForName(currentDisplayName);
        }

        var newCustomName = PromptForName(currentCustomName);
        if (string.IsNullOrWhiteSpace(newCustomName))
        {
            return null;
        }

        return BuildDbDisplayName(dbNumber, newCustomName.Trim());
    }

    private string? PromptForDbBlockProperties(ProjectTreeNode dbNode)
    {
        var currentDisplayName = dbNode.Name;
        if (!TryParseDbDisplayName(currentDisplayName, out var currentCustomName, out var currentDbNumber))
        {
            MessageBox.Show(this, $"无法解析数据块编号：{currentDisplayName}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }

        var dbFolder = dbNode.Parent;
        if (dbFolder is null)
        {
            MessageBox.Show(this, "未找到数据块父节点。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }

        var nameInput = new TextBox
        {
            Text = currentCustomName,
            MinWidth = 360,
            Margin = new Thickness(0, 8, 0, 10)
        };
        var namespaceInput = new TextBox
        {
            Text = string.Empty,
            MinWidth = 360,
            Margin = new Thickness(0, 0, 0, 10),
            IsReadOnly = true
        };
        var typeInput = new TextBox
        {
            Text = "DB",
            MinWidth = 120,
            Margin = new Thickness(0, 0, 0, 10),
            IsReadOnly = true
        };
        var typeDescInput = new TextBox
        {
            Text = "全局 DB",
            MinWidth = 120,
            Margin = new Thickness(8, 0, 0, 10),
            IsReadOnly = true
        };
        var languageBox = new ComboBox
        {
            MinWidth = 360,
            Margin = new Thickness(0, 0, 0, 10),
            IsEnabled = false
        };
        languageBox.Items.Add("DB");
        languageBox.SelectedIndex = 0;

        var numberInput = new TextBox
        {
            MinWidth = 180,
            Margin = new Thickness(0, 0, 0, 4),
            IsEnabled = false
        };

        var manualRadio = new RadioButton
        {
            Content = "手动",
            Margin = new Thickness(0, 0, 0, 4)
        };
        var autoRadio = new RadioButton
        {
            Content = "自动",
            IsChecked = true,
            Margin = new Thickness(0, 0, 0, 10)
        };
        var validationText = new TextBlock
        {
            Foreground = GetThemeBrush("DangerBrush", Brushes.Firebrick),
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };

        var form = new Grid
        {
            Margin = new Thickness(0, 8, 0, 12)
        };
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var nameLabel = new TextBlock { Text = "名称：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var namespaceLabel = new TextBlock { Text = "命名空间：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var typeLabel = new TextBlock { Text = "类型：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var languageLabel = new TextBlock { Text = "语言：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var numberLabel = new TextBlock { Text = "编号：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Top };

        Grid.SetRow(nameLabel, 0);
        Grid.SetColumn(nameLabel, 0);
        Grid.SetRow(nameInput, 0);
        Grid.SetColumn(nameInput, 1);

        Grid.SetRow(namespaceLabel, 1);
        Grid.SetColumn(namespaceLabel, 0);
        Grid.SetRow(namespaceInput, 1);
        Grid.SetColumn(namespaceInput, 1);

        var typePanel = new StackPanel { Orientation = Orientation.Horizontal };
        typePanel.Children.Add(typeInput);
        typePanel.Children.Add(typeDescInput);
        Grid.SetRow(typeLabel, 2);
        Grid.SetColumn(typeLabel, 0);
        Grid.SetRow(typePanel, 2);
        Grid.SetColumn(typePanel, 1);

        Grid.SetRow(languageLabel, 3);
        Grid.SetColumn(languageLabel, 0);
        Grid.SetRow(languageBox, 3);
        Grid.SetColumn(languageBox, 1);

        var numberPanel = new StackPanel();
        numberPanel.Children.Add(numberInput);
        numberPanel.Children.Add(manualRadio);
        numberPanel.Children.Add(autoRadio);
        Grid.SetRow(numberLabel, 4);
        Grid.SetColumn(numberLabel, 0);
        Grid.SetRow(numberPanel, 4);
        Grid.SetColumn(numberPanel, 1);

        form.Children.Add(nameLabel);
        form.Children.Add(nameInput);
        form.Children.Add(namespaceLabel);
        form.Children.Add(namespaceInput);
        form.Children.Add(typeLabel);
        form.Children.Add(typePanel);
        form.Children.Add(languageLabel);
        form.Children.Add(languageBox);
        form.Children.Add(numberLabel);
        form.Children.Add(numberPanel);

        var okButton = new Button
        {
            Content = "确定",
            Width = 84,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 84,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel
        {
            Margin = new Thickness(16)
        };
        root.Children.Add(form);
        root.Children.Add(validationText);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = currentDisplayName,
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        void SetValidation(string? message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                validationText.Text = string.Empty;
                validationText.Visibility = Visibility.Collapsed;
                numberInput.ClearValue(BackgroundProperty);
                okButton.IsEnabled = true;
                return;
            }

            validationText.Text = message;
            validationText.Visibility = Visibility.Visible;
            numberInput.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEAEA"));
            okButton.IsEnabled = false;
        }

        void ValidateNumber()
        {
            if (manualRadio.IsChecked != true)
            {
                SetValidation(null);
                return;
            }

            if (!int.TryParse(numberInput.Text.Trim(), out var manualDbNumber) || manualDbNumber <= 0)
            {
                SetValidation("编号必须是大于 0 的整数。");
                return;
            }

            if (IsDbNumberInUse(dbFolder, manualDbNumber, dbNode))
            {
                SetValidation($"DB{manualDbNumber} 已存在，请更换编号。");
                return;
            }

            SetValidation(null);
        }

        autoRadio.Checked += (_, _) =>
        {
            numberInput.IsEnabled = false;
            numberInput.Text = FindNextDbNumber(dbFolder, excludeDbNumber: currentDbNumber).ToString(CultureInfo.InvariantCulture);
            ValidateNumber();
        };
        manualRadio.Checked += (_, _) =>
        {
            numberInput.IsEnabled = true;
            numberInput.Text = currentDbNumber.ToString(CultureInfo.InvariantCulture);
            numberInput.Focus();
            numberInput.SelectAll();
            ValidateNumber();
        };
        numberInput.TextChanged += (_, _) => ValidateNumber();

        numberInput.Text = FindNextDbNumber(dbFolder, excludeDbNumber: currentDbNumber).ToString(CultureInfo.InvariantCulture);

        string? resultName = null;
        okButton.Click += (_, _) =>
        {
            var customName = nameInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(customName))
            {
                MessageBox.Show(dialog, "名称不能为空。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int targetDbNumber;
            if (autoRadio.IsChecked == true)
            {
                targetDbNumber = FindNextDbNumber(dbFolder, excludeDbNumber: currentDbNumber);
            }
            else
            {
                if (!int.TryParse(numberInput.Text.Trim(), out targetDbNumber) || targetDbNumber <= 0)
                {
                    MessageBox.Show(dialog, "编号必须是大于 0 的整数。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (IsDbNumberInUse(dbFolder, targetDbNumber, dbNode))
                {
                    MessageBox.Show(dialog, $"DB{targetDbNumber} 已存在，请更换编号。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            resultName = BuildDbDisplayName(targetDbNumber, customName);
            dialog.DialogResult = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) =>
        {
            dialog.DialogResult = false;
            dialog.Close();
        };

        nameInput.Focus();
        nameInput.SelectAll();

        var result = dialog.ShowDialog();
        return result == true ? resultName : null;
    }

    private string? PromptForDbBlockName(ProjectTreeNode dbFolder, int nextDbNumber)
    {
        var nameInput = new TextBox
        {
            Text = $"数据块_{nextDbNumber}",
            MinWidth = 320,
            Margin = new Thickness(0, 8, 0, 10)
        };

        var numberInput = new TextBox
        {
            Text = nextDbNumber.ToString(CultureInfo.InvariantCulture),
            MinWidth = 120,
            IsEnabled = false
        };

        var autoRadio = new RadioButton
        {
            Content = "自动",
            IsChecked = true,
            Margin = new Thickness(0, 8, 12, 0)
        };
        var manualRadio = new RadioButton
        {
            Content = "手动",
            Margin = new Thickness(0, 8, 0, 0)
        };

        autoRadio.Checked += (_, _) =>
        {
            numberInput.IsEnabled = false;
            numberInput.Text = FindNextDbNumber(dbFolder).ToString(CultureInfo.InvariantCulture);
        };

        manualRadio.Checked += (_, _) =>
        {
            numberInput.IsEnabled = true;
            numberInput.Focus();
            numberInput.SelectAll();
        };

        var modePanel = new StackPanel
        {
            Orientation = Orientation.Horizontal
        };
        modePanel.Children.Add(autoRadio);
        modePanel.Children.Add(manualRadio);

        var form = new Grid
        {
            Margin = new Thickness(0, 8, 0, 12)
        };
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var nameLabel = new TextBlock { Text = "名称：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var numberLabel = new TextBlock { Text = "编号：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        var modeLabel = new TextBlock { Text = "方式：", Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };

        Grid.SetRow(nameLabel, 0);
        Grid.SetColumn(nameLabel, 0);
        Grid.SetRow(nameInput, 0);
        Grid.SetColumn(nameInput, 1);
        Grid.SetRow(numberLabel, 1);
        Grid.SetColumn(numberLabel, 0);
        Grid.SetRow(numberInput, 1);
        Grid.SetColumn(numberInput, 1);
        Grid.SetRow(modeLabel, 2);
        Grid.SetColumn(modeLabel, 0);
        Grid.SetRow(modePanel, 2);
        Grid.SetColumn(modePanel, 1);

        form.Children.Add(nameLabel);
        form.Children.Add(nameInput);
        form.Children.Add(numberLabel);
        form.Children.Add(numberInput);
        form.Children.Add(modeLabel);
        form.Children.Add(modePanel);

        var description = new TextBlock
        {
            Text = "描述：数据块（DB）保存程序数据。",
            Foreground = GetThemeBrush("HintBrush", Brushes.DimGray),
            Margin = new Thickness(0, 0, 0, 14)
        };

        var okButton = new Button
        {
            Content = "确定",
            Width = 84,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 84,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel
        {
            Margin = new Thickness(16)
        };
        root.Children.Add(new TextBlock
        {
            Text = "添加数据块",
            FontWeight = FontWeights.SemiBold,
            FontSize = 14
        });
        root.Children.Add(form);
        root.Children.Add(description);
        root.Children.Add(buttonPanel);

        var dialog = new Window
        {
            Title = "添加数据块",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        string? resultName = null;
        okButton.Click += (_, _) =>
        {
            var customName = nameInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(customName))
            {
                MessageBox.Show(dialog, "名称不能为空。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dbNumberText = autoRadio.IsChecked == true
                ? FindNextDbNumber(dbFolder).ToString(CultureInfo.InvariantCulture)
                : numberInput.Text.Trim();

            if (!int.TryParse(dbNumberText, out var dbNumber) || dbNumber <= 0)
            {
                MessageBox.Show(dialog, "编号必须是大于 0 的整数。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (IsDbNumberInUse(dbFolder, dbNumber))
            {
                MessageBox.Show(dialog, $"DB{dbNumber} 已存在，请更换编号。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            resultName = BuildDbDisplayName(dbNumber, customName);
            dialog.DialogResult = true;
            dialog.Close();
        };

        cancelButton.Click += (_, _) =>
        {
            dialog.DialogResult = false;
            dialog.Close();
        };

        nameInput.Focus();
        nameInput.SelectAll();

        var result = dialog.ShowDialog();
        return result == true ? resultName : null;
    }

    private string? PromptForValue(VariableRow row)
    {
        var textInput = new TextBox
        {
            Text = row.Value,
            MinWidth = 260,
            Margin = new Thickness(0, 10, 0, 12),
            Visibility = row.Type == PlcValueType.Bool ? Visibility.Collapsed : Visibility.Visible
        };

        var boolInput = new ComboBox
        {
            MinWidth = 260,
            Margin = new Thickness(0, 10, 0, 12),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            Visibility = row.Type == PlcValueType.Bool ? Visibility.Visible : Visibility.Collapsed
        };
        boolInput.Items.Add("TRUE");
        boolInput.Items.Add("FALSE");
        var currentBool = string.Equals(row.Value.Trim(), "true", StringComparison.OrdinalIgnoreCase) ||
                          row.Value.Trim() == "1";
        boolInput.SelectedItem = currentBool ? "TRUE" : "FALSE";

        var okButton = new Button
        {
            Content = "确定",
            Width = 80,
            IsDefault = true
        };
        var cancelButton = new Button
        {
            Content = "取消",
            Width = 80,
            IsCancel = true,
            Margin = new Thickness(8, 0, 0, 0)
        };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttons.Children.Add(okButton);
        buttons.Children.Add(cancelButton);

        var root = new StackPanel
        {
            Margin = new Thickness(16)
        };
        root.Children.Add(new TextBlock { Text = $"请输入 {row.Name} 的新监视值：" });
        root.Children.Add(textInput);
        root.Children.Add(boolInput);
        root.Children.Add(buttons);

        var dialog = new Window
        {
            Title = "修改监视值",
            Content = root,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            ShowInTaskbar = false
        };
        ApplyHtmlDialogTheme(dialog);

        okButton.Click += (_, _) =>
        {
            dialog.DialogResult = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) =>
        {
            dialog.DialogResult = false;
            dialog.Close();
        };

        if (row.Type == PlcValueType.Bool)
        {
            boolInput.Focus();
        }
        else
        {
            textInput.Focus();
            textInput.SelectAll();
        }

        var result = dialog.ShowDialog();
        if (result != true)
        {
            return null;
        }

        return row.Type == PlcValueType.Bool
            ? (boolInput.SelectedItem?.ToString() ?? "FALSE")
            : textInput.Text;
    }

    private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child is not null)
        {
            if (child is T target)
            {
                return target;
            }

            child = VisualTreeHelper.GetParent(child);
        }

        return null;
    }
}

public sealed record PlcTypeOption(PlcValueType Value, string Label);

public sealed record PlcEndpointSettings(string Address, int Port, bool AutoStart);

public sealed record ArrayTypeDefinition(PlcValueType ElementType, int LowerBound, int UpperBound, int? StringLength);

public sealed record ArrayElementTypeOption(PlcValueType Value, string Label, bool RequiresLength = false);

public sealed class SimulatorState
{
    public List<ProjectNodeState> ProjectNodes { get; set; } = [];
    public List<VariableRowState> Variables { get; set; } = [];
    public Dictionary<int, byte[]> DbBlocks { get; set; } = [];
    public Dictionary<string, PlcEndpointSettingState> PlcSettings { get; set; } = [];
}

public sealed class ProjectNodeState
{
    public string Name { get; set; } = string.Empty;
    public ProjectNodeType NodeType { get; set; }
    public List<ProjectNodeState> Children { get; set; } = [];
}

public sealed class VariableRowState
{
    public string Name { get; set; } = string.Empty;
    public int DbNumber { get; set; }
    public int Offset { get; set; }
    public int Bit { get; set; }
    public PlcValueType Type { get; set; }
    public int StringLength { get; set; } = 256;
    public PlcValueType ArrayElementType { get; set; } = PlcValueType.Bool;
    public int ArrayLowerBound { get; set; }
    public int ArrayUpperBound { get; set; } = 1;
    public string StartValue { get; set; } = "FALSE";
    public string Value { get; set; } = "FALSE";
}

public sealed class PlcEndpointSettingState
{
    public string Address { get; set; } = "";
    public int Port { get; set; } = 102;
    public bool AutoStart { get; set; } = true;
}
