using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace PdfToDocx
{
    public class PdfEnvelopItem
    {
        public int Page { get; set; }

        public string Sequence { get => SequenceWord.Text.Substring(1, 4); }

        public string Owner
        {
            get
            {
                string o = OwnerWord.Text.Length > 5 ? OwnerWord.Text.Substring(5, 1) : "  ";
                return IdWord.Text.Substring(6, 1).Equals("1") ? o + "先生" : o + "小姐";
            }
        }

        public string Zip { get; set; }

        public string Address { get => AddressWord != null ? AddressWord.Text.Substring(2) : null; }


        public AddressImageItem AddressImage;


        [JsonIgnore]
        public Word SequenceWord { get; set; }

        [JsonIgnore]
        public Word AddressWord { get; set; }

        [JsonIgnore]
        public Word OtherWord { get; set; }

        [JsonIgnore]
        public PdfPoint SequencePoint { get => SequenceWord.Letters[0].Location; }

        [JsonIgnore]
        public PdfPoint IdPoint { get => IdWord.Letters[0].Location; }

        [JsonIgnore]
        public PdfPoint OtherWordPoint { get => OtherWord.Letters[0].Location; }

        [JsonIgnore]
        public Word OwnerWord { get; set; }

        [JsonIgnore]
        public Word IdWord { get; set; }



    }
}
