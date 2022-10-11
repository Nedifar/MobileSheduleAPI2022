using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ApiShedule2022.Models
{
    public class FloorCabinet
    {
        public string Name { get; set; }
        public List<DayWeekClass> DayWeeks { get; set; } = new List<DayWeekClass>();
    }
}
