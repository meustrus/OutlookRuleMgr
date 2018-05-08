using System.IO;
using Newtonsoft.Json;

namespace OutlookRuleMgr.Utilities
{
    public static class Json
    {
        public static T ReadFile<T>(string filename)
        {
            using (var fileStream = File.OpenText(filename))
            using (var jsonReader = new JsonTextReader(fileStream))
            {
                return CreateJsonSerializer().Deserialize<T>(jsonReader);
            }
        }

        public static void WriteFile(object obj, string filename)
        {
            using (var file = File.CreateText(filename))
            {
                CreateJsonSerializer().Serialize(file, obj);
            }
        }

        private static JsonSerializer CreateJsonSerializer()
        {
            var jsonSerializer = JsonSerializer.Create(new JsonSerializerSettings
            {
                DefaultValueHandling = DefaultValueHandling.Ignore
            });
            jsonSerializer.Formatting = Formatting.Indented;
            return jsonSerializer;
        }
    }
}
