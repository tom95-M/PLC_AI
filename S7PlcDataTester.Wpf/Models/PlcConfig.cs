namespace S7PlcDataTester.Wpf.Models
{
    public enum PlcType
    {
        SiemensS7,
        Mitsubishi,
        Inovance,
        KeyenceKV
    }

    public class PlcConfig
    {
        public PlcType Type { get; set; }
        public string IpAddress { get; set; }
        public int Port { get; set; }
        public string Rack { get; set; }
        public string Slot { get; set; }
        public string Station { get; set; }
        public string Protocol { get; set; }
        public string Name { get; set; }

        public PlcConfig()
        {
            IpAddress = "127.0.0.1";
            Port = 102;
            Rack = "0";
            Slot = "1";
            Station = "0";
            Protocol = "TCP";
            Name = "PLC";
        }
    }
}
