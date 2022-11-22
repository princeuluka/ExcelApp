using ExcelApp.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using IronXL;
using System.Data;
using System.Diagnostics;
using System.Text;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using Aspose.Cells;
using ExcelApp.Service;
using NPOI.Util;

namespace ExcelApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IHostingEnvironment Environment;
        private readonly ApplicationDbContext _dbContext;

        public HomeController(ILogger<HomeController> logger, IHostingEnvironment _environment, ApplicationDbContext dbContext)
        {
            _logger = logger;
            Environment = _environment;
            _dbContext = dbContext;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet]
        public IActionResult UploadForm()
        {
            var data = _dbContext.Excel.ToList();
            return View(data);
        }
        [HttpPost]
        public  DataTable Import()
        {
            UploadService upload = new UploadService();
            IFormFile file = Request.Form.Files[0];
            string folderName = "UploadExcel";
            string webRootPath = Environment.WebRootPath;
            string newPath = Path.Combine(webRootPath, folderName);
            StringBuilder sb = new StringBuilder();
            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }
            if (file.Length > 0)
            {
                string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                ISheet sheet;
                string fullPath = Path.Combine(newPath, file.FileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                    stream.Position = 0;
                }

                var importer = upload.Mappings = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Identity","Identity"),
                    new KeyValuePair<string, string>("FirstName","FirstName"),
                    new KeyValuePair<string, string>("Surname","Surname"),
                    new KeyValuePair<string, string>("Age","Age"),
                    new KeyValuePair<string, string>("Sex","Sex"),
                    new KeyValuePair<string, string>("Mobile","Mobile"),
                    new KeyValuePair<string, string>("Active","Active")
                };

                var list = upload.Import<ExcelModel>(fullPath);

                ExcelModel model = new ExcelModel();
                DataTable table = upload.ToDataTable<ExcelModel>(list);

                foreach (DataRow item in table.Rows)
                {
                    model.Identity = (int)item[0];
                    model.FirstName = (string)item[1];
                    model.Surname = (string)item[2];
                    model.Age = (string)item[3];
                    model.Sex = (string)item[4];
                    model.Active = (string)item[5];

                    AddData(model);
                }
                return table;
            }
            else
            {
                return null;
            }
            
        }

        public ActionResult Download()
        {
            string Files = "wwwroot/UploadExcel/";
            byte[] fileBytes = System.IO.File.ReadAllBytes(Files);
            System.IO.File.WriteAllBytes(Files, fileBytes);
            MemoryStream ms = new MemoryStream(fileBytes);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "employee.xlsx");
        }

        

        public IActionResult UpdateData(ExcelModel model) 
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            _dbContext.Excel.Update(model);
            _dbContext.SaveChanges();
            return View(nameof(UpdateData));
        }

        public void AddData(ExcelModel model)
        {
            _dbContext.Excel.Add(model);
            _dbContext.SaveChanges();
        }
    }
}