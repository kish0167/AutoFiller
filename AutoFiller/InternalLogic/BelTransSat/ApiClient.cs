using System.Net.Http;
using System.Text.Json;
using DataTracker.Utility;

namespace ExcelParser.BelTransSat;

enum RequestType
{
    OBJECT_STAT_DATA = 0
}

public class ApiClient
{
    private HttpClient _httpClient;
    private string _token;
    private List<Response> _responsesPool;


    public ApiClient(string token)
    {
        _httpClient = new HttpClient();
        _token = token;
        _responsesPool = new List<Response>();
    }

    public async Task<RootObject> GetVehiclesInfo(DateTime date)
    {
        foreach (Response oldResponce in _responsesPool)
        {
            if (oldResponce.ResponseDate == date)
            {
                //Logger.Log("\tAPI Request avoided - response already exists");
                return oldResponce.ResponseObject;
            }
        }

        Logger.Log("\tCalling BTS...");
        string URL = GenerateURL(date, RequestType.OBJECT_STAT_DATA);
        HttpResponseMessage response = await _httpClient.GetAsync(URL);

        if (response.IsSuccessStatusCode)
        {
            Logger.Log("\tResponse acquired for date: " + date.ToString("yyyy-MM-dd"));
            string jsonResponse = await response.Content.ReadAsStringAsync();
            RootObject rootObject = JsonSerializer.Deserialize<RootObject>(jsonResponse);
            _responsesPool.Add(new Response(date, rootObject));
            return rootObject;
        }
        else
        {
            throw new HttpRequestException("Failed to fetch tracking data.");
        }
    }

    private string GenerateURL(DateTime date, RequestType requestType)
    {
        string begin = "https://api.nav.by/info/integration.php?type=";
        string type = requestType + "&token=";
        string token = _token + "&from=";
        string from = date.ToString("yyyy-MM-dd") + "%2000%3A00%3A00&to=";
        string to = date.ToString("yyyy-MM-dd") + "%2023%3A59%3A59";
        string url = begin + type + token + from + to + "&output=json";
        return url;
    }


    private class Response
    {
        public DateTime ResponseDate;
        public RootObject ResponseObject;

        public Response(DateTime d, RootObject r)
        {
            ResponseDate = d;
            ResponseObject = r;
        }
    }
}