using NAudio.CoreAudioApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XCSoftware
{
    public class XCDeviceChannel
    {
        public int ID { get; set; }

        public string Name { get; set; }

        public int Value { get; set; }

        public MMDevice MMDevice { get; set; }

        public string MMDeviceID { get; set; }
    }
}
