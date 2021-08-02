using System.Collections.Generic;

namespace ConsoleApp.Models
{
    public class Topic
    {
        public string SessionName { get; set; }
        public string Title { get; set; }
        public string PresentationAbstract { get; set; }
        public List<Participant> Participants { get; set; } = new List<Participant>();
    }
}
