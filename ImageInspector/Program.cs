using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using UnitsNet;

namespace ImageInspector
{
    class Program
    {
        static void Main(string[] args)
        {
            var imagePath = ConfigurationManager.AppSettings["ImagePath"];
            var csvPath = ConfigurationManager.AppSettings["CsvPath"];
            var searchPattern = "*.jpeg";

            var records = from f in Directory.EnumerateFiles(imagePath, searchPattern)
                          select ImageData.From(f);

            using (var writer = new StreamWriter(
                Path.Combine(csvPath, $"{DateTime.Now.ToString("yyyyMMddhhmmssfff")}.csv"),
                false,
                System.Text.Encoding.GetEncoding("shift_jis")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Configuration.RegisterClassMap<ImageMapper>();
                csv.WriteRecords(records);
            }
        }
    }

    class ImageData
    {
        public string FileName { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string PixelFormat { get; set; }

        private ImageData()
        { }

        public static ImageData From(string path)
        {
            var fileName = new FileInfo(path).Name;
            using (System.Drawing.Image img = System.Drawing.Image.FromFile(path))
            {
                return new ImageData
                {
                    FileName = fileName,
                    Width = Length.FromInches(img.Width / img.HorizontalResolution).Centimeters,
                    Height = Length.FromInches(img.Height / img.VerticalResolution).Centimeters,
                    PixelFormat = img.PixelFormat.ToString(),
                };
            }
        }
    }

    class ImageMapper : ClassMap<ImageData>
    {
        public ImageMapper()
        {
            Map(x => x.FileName).Index(0).Name("ファイル名");
            Map(x => x.Width).Index(1).Name("幅");
            Map(x => x.Height).Index(2).Name("高さ");
            Map(x => x.PixelFormat).Index(3);
        }
    }
}
