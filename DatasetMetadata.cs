using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Newtonsoft.Json;

namespace PptExcelSync
{
    public class DatasetMetadata
    {
        public string UploadedBy { get; set; }
        public DateTime UploadedAt { get; set; }
        public List<CalculatedFieldInfo> CalculatedFields { get; set; } = new List<CalculatedFieldInfo>();

        public static DatasetMetadata Load(string datasetPath)
        {
            string metaJson = datasetPath + ".meta.json";
            string metaTxt = datasetPath + ".meta.txt";

            // ✅ Case 1: JSON exists → load it
            if (File.Exists(metaJson))
            {
                return JsonConvert.DeserializeObject<DatasetMetadata>(
                    File.ReadAllText(metaJson)
                );
            }

            // ✅ Case 2: Migrate TXT → JSON
            if (File.Exists(metaTxt))
            {
                var lines = File.ReadAllLines(metaTxt);
                var metadata = new DatasetMetadata();

                foreach (var line in lines)
                {
                    if (line.StartsWith("uploadedBy="))
                        metadata.UploadedBy = line.Substring("uploadedBy=".Length);

                    if (line.StartsWith("uploadedAt=") &&
                        DateTime.TryParse(line.Substring("uploadedAt=".Length), out var dt))
                        metadata.UploadedAt = dt;
                }

                // Ensure CalculatedFields is not null
                metadata.CalculatedFields = new List<CalculatedFieldInfo>();

                // Save immediately as JSON so next time we don’t need TXT anymore
                metadata.Save(datasetPath);

                // Optionally: delete old .meta.txt to keep folder clean
                // File.Delete(metaTxt);

                return metadata;
            }

            // ✅ Case 3: Nothing exists → create new
            return new DatasetMetadata
            {
                UploadedBy = Environment.UserName,
                UploadedAt = DateTime.UtcNow,
                CalculatedFields = new List<CalculatedFieldInfo>()
            };
        }

        public void Save(string datasetPath)
        {
            string metaFile = datasetPath + ".meta.json";
            File.WriteAllText(metaFile, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }

    public class CalculatedFieldInfo
    {
        public string FieldName { get; set; }
        public string Formula { get; set; }
    }

    public class ConditionalRule
    {
        public string Field { get; set; }
        public string Operator { get; set; } // >, <, >=, <=, =
        public double Threshold { get; set; }
        public Color Color { get; set; }

    }


}
