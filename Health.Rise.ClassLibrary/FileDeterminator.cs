using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Health.Rise.ClassLibrary
{
    public static class FileDeterminator
    {
        private static readonly Dictionary<string, byte[]> Formats = new Dictionary<string, byte[]>()
        {
            {"PNG", new byte[]{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52}},
            {"JPG", new byte[]{0xFF, 0xD8, 0xFF}},
            {"GIF", new byte[]{0x47, 0x49, 0x46, 0x38}},
            {"DOCorXLS", new byte[] {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }},
            {"PDF", new byte[] { 0x25, 0x50, 0x44, 0x46} },
            {"DOCXorXLSX", new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06,0x00} }
        };
       

        public static ICollection<string> SupportedFormats
        {
            get { return Formats.Keys; }
        }

        public static byte[] GetPrefix(string format)
        {
            return Formats[format];
        }

        public static string DetermineFormat(Stream stream)
        {
            stream.Position = 0;
            int length = (int)Math.Min(stream.Length, Formats.Max(f => f.Value.Length));
            byte[] bytes = new byte[length];
            stream.Read(bytes, 0, length);
            stream.Position = 0;
            foreach (var format in Formats)
            {
                if (HasPrefix(bytes, format.Value))
                    return format.Key;
            }
            return null;
        }

        private static bool HasPrefix(byte[] file, byte[] prefix)
        {
            if (prefix.Length > file.Length)
                return false;
            for (int i = 0; i < prefix.Length; i++)
            {
                if (prefix[i] != file[i])
                    return false;
            }
            return true;
        }
    }
}
