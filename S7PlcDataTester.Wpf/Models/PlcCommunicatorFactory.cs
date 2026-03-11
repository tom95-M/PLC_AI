using S7PlcDataTester.Wpf.Interfaces;
using S7PlcDataTester.Wpf.Communicators;

namespace S7PlcDataTester.Wpf.Models
{
    public static class PlcCommunicatorFactory
    {
        public static IPlcCommunicator CreateCommunicator(PlcConfig config)
        {
            switch (config.Type)
            {
                case PlcType.SiemensS7:
                    return new SiemensS7PlcCommunicator(config);
                case PlcType.Mitsubishi:
                    return new MitsubishiPlcCommunicator(config);
                case PlcType.Inovance:
                    return new InovancePlcCommunicator(config);
                case PlcType.KeyenceKV:
                    return new KeyenceKvPlcCommunicator(config);
                default:
                    throw new System.ArgumentException($"Unknown PLC type: {config.Type}");
            }
        }
    }
}
