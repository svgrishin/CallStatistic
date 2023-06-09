namespace CallStatistic.Models
{
    public class Calls
    {
        public int id { get; set; }
        public DateTime dateOf { get; set; }
        public string fromPhone { get; set; } = string.Empty;
        public string toPhone { get; set; } = string.Empty;
        public string duration { get; set; }

        public Calls() { }
        public Calls(DateTime dateOf, string fromPhone, string toPhone, string duration)
        {
            this.dateOf = dateOf;
            this.fromPhone = fromPhone;
            this.toPhone = toPhone;
            this.duration = duration;
        }
    }
}