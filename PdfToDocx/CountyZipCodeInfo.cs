using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace PdfToDocx
{

    public class CountyZipCodeInfo
    {
        [JsonPropertyName("districts")]
        public List<District> Districts { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }
    }


    public class District
    {
        [JsonPropertyName("zip")]
        public string Zip { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }
    }

}
