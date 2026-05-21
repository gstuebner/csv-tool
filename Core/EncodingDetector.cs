using System;
using System.IO;
using System.Text;

namespace CsvTool.Core
{
    public static class EncodingDetector
    {
        public static Encoding Detect(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length >= 3)
                {
                    byte[] bom = new byte[3];
                    stream.Read(bom, 0, 3);
                    if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
                }
            }

            byte[] buffer = new byte[4096];
            using (var stream = File.OpenRead(path))
            {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (IsUtf8(buffer, read)) return Encoding.UTF8;
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(1252);
        }

        private static bool IsUtf8(byte[] buffer, int length)
        {
            int i = 0;
            while (i < length)
            {
                byte c = buffer[i];
                if (c < 0x80) i++;
                else if ((c & 0xE0) == 0xC0) { if (i + 1 >= length || (buffer[i + 1] & 0xC0) != 0x80) return false; i += 2; }
                else if ((c & 0xF0) == 0xE0) { if (i + 2 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80) return false; i += 3; }
                else if ((c & 0xF8) == 0xF0) { if (i + 3 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80 || (buffer[i + 3] & 0xC0) != 0x80) return false; i += 4; }
                else return false;
            }
            return true;
        }
    }
}
