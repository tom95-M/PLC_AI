using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace S7PlcSimulator;

public enum ProjectNodeType
{
    PlcRoot,
    ProgramFolder,
    ProgramGroup,
    DataBlockFolder,
    DataBlock
}

public sealed class ProjectTreeNode : INotifyPropertyChanged
{
    private string _name = string.Empty;

    public string Name
    {
        get => _name;
        set
        {
            if (_name == value)
            {
                return;
            }

            _name = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
        }
    }

    public ProjectNodeType NodeType { get; init; }
    public ObservableCollection<ProjectTreeNode> Children { get; } = new();
    public ProjectTreeNode? Parent { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;
}
