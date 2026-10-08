using System.Web.Script.Serialization;

namespace RemoteServiceManager.Services
{
    public static class JsonHelper
    {
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer
        {
            MaxJsonLength = int.MaxValue
        };

        public static string Serialize<T>(T obj)
        {
            return Serializer.Serialize(obj);
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return default(T);

            return Serializer.Deserialize<T>(json);
        }
    }
}
