using System.Data;
using System.IO;
using System.Reflection;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;

namespace ExcelApp.Service
{
    public class UploadService
    {
        private IHostingEnvironment Environment;


        public UploadService(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }

        public UploadService()
        {
        }

        public List<KeyValuePair<string, string>> Mappings;

        public void Upload(IFormFile postedFile)
        {
            if (postedFile != null)
            {
                //Create a Folder.
                string path = Path.Combine(this.Environment.WebRootPath, "Uploads");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                //Save the uploaded Excel file.
                string fileName = Path.GetFileName(postedFile.FileName);
                string filePath = Path.Combine(path, fileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    postedFile.CopyTo(stream);
                }
            }
        }

      
        public List <T> Import<T>(string file)
        {
            List<T> list = new List<T>();
            List<string>lines = File.ReadAllLines(file).ToList();
            string headerLine = lines[0];
            var headerInfo = headerLine.Split(',').ToList().Select((v, i) => new { ColName = v, ColIndex = i });

            Type type = typeof(T);
            var properties = type.GetProperties();

            var dataLines = lines.Skip(1);

            dataLines.ToList().ForEach(line =>
            {
                var values = line.Split(',');
                T obj = (T)Activator.CreateInstance(type);

                foreach (var prop in properties)
                {
                    var mapping = Mappings.SingleOrDefault(m => m.Value == prop.Name);
                    var colName = mapping.Key;
                    var colIndex = headerInfo.SingleOrDefault(s => s.ColName == colName).ColIndex;
                    var value = values[colIndex];
                    var propType = prop.PropertyType;
                    prop.SetValue(obj,Convert.ChangeType(value,propType));
                }
                list.Add(obj);
            });
            return list;
        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }


         
    }
}
