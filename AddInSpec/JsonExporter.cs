using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace AddInSpec
{
    internal class JsonExporter
    {
        public void ExportAssemblyToJson(Assembly assembly, string filePath)
        {
            if (assembly == null)
            {
                MessageBox.Show("Assembly information cannot be null.");
               
            }

            var root = new Root { Assembly = assembly };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var json = JsonSerializer.Serialize(root, options);
            File.WriteAllText(filePath, json);
        }
    }
}
