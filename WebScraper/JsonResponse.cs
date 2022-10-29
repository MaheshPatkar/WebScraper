namespace WebScraper
{
    public class GetJsonResponse
    {
        public bool Success { get; set; }

        public dynamic? Data { get; set; }
    }


    public class PostJsonResponse
    {
        public bool Success { get; set; }

        public string? Costs { get; set; }
    }


}
