using System.Collections.Generic;

namespace ApiShedule2022.Models
{
    public class ListPack
    {
        public List<string> Groups { get; set; } = new List<string>();
        public List<string> Cabinets { get; set; } = new List<string>();
        public List<string> Teachers { get; set; } = new List<string>();
    }
}
