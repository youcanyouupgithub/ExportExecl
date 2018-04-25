using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

public class JsonHelper
{
    public static string JsonSerializer<T>(T t)
    {
        DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(T));
        System.IO.MemoryStream memoryStream = new System.IO.MemoryStream();
        dataContractJsonSerializer.WriteObject(memoryStream, t);
        string @string = System.Text.Encoding.UTF8.GetString(memoryStream.ToArray());
        memoryStream.Close();
        return @string;
    }
    public static T JsonDeserialize<T>(string jsonString)
    {
        DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(T));
        System.IO.MemoryStream stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(jsonString));
        return (T)((object)dataContractJsonSerializer.ReadObject(stream));
    }
}
