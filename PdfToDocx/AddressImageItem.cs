using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;
using UglyToad.PdfPig.Content;

namespace PdfToDocx
{
    public class AddressImageItem
    {
        public int Page { get; set; }

        public string FileName { get; set; }

        [JsonIgnore]
        public IPdfImage AddressImage { get; set; }

    }
}
