using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XCSoftware
{
    class XCDevice
    {
        public int ID { get; set; }

        public string Name { get; set; }

        public string SerialNumber { get; set; }

        public List<XCDeviceChannel> XCDeviceChannels { get; set; }

        public XCDevice()
        {
            this.XCDeviceChannels = new List<XCDeviceChannel>();
        }
    }
}
