using System.Globalization;
using System.Windows.Data;

namespace S7PlcSimulator;

public sealed class ProjectNodeTypeToEmojiConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not ProjectNodeType nodeType)
        {
            return "📁";
        }

        return nodeType switch
        {
            ProjectNodeType.PlcRoot => "🎛️",
            ProjectNodeType.ProgramFolder or ProjectNodeType.ProgramGroup => "📄",
            ProjectNodeType.DataBlockFolder or ProjectNodeType.DataBlock => "🛢️",
            _ => "📁"
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
