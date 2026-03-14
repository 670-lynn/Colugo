namespace Colugo.Models
{
    public class ColugoModel
    {
        public string Input { get; set; }

        public string Process()
        {
            return $"Colugo: {Input}";
        }
    }
}
