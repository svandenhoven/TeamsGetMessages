using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeamsGetMessages
{
    public class Settings
    {
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string TeamId { get; set; }
        public string ChannelId { get; set; }
        public string NotificationUrl { get; set; }

    }
}
