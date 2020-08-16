using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Net.Http;
using Giveaway.Helper;
using MailMerge = Giveaway.Models.MailMerge;
using Giveaway.Infra;

namespace Giveaway.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> logger;
        private readonly IWordMailMerge wordMailMerge;

        public HomeController(ILogger<HomeController> _logger,IWordMailMerge _wordMailMerge)
        {
            logger = _logger;
            wordMailMerge = _wordMailMerge;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Index(MailMerge mailMergeFiles)
        {
            //Call uploadfile function to upload files and generate email
            if (ModelState.IsValid)
            {
                await UploadFile(mailMergeFiles);
                TempData["msg"] = "Mail Merge Completed Successfully.";
            }
            return View();
        }
        /// <summary>
        /// Main fucntion to upload file on server and send email
        /// </summary>
        /// <param name="mailMergeFiles"></param>
        /// <returns></returns>
        public async Task<bool> UploadFile(MailMerge mailMergeFiles)
        {
            string path = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "Upload")); ;
            string templateFileName="", valuesFileName="";

            try
            {
                MultipartFormDataContent formDataContent = new MultipartFormDataContent();

                if (mailMergeFiles.TemplateFile != null && mailMergeFiles.TemplateFile.Length > 0
                    && mailMergeFiles.ValuesFile != null && mailMergeFiles.ValuesFile.Length > 0)
                {
                    //Upload template file to server
                    templateFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + mailMergeFiles.TemplateFile.FileName;
                    
                    using (var filestream = new FileStream(Path.Combine(path, templateFileName), FileMode.Create))
                    {
                        await mailMergeFiles.TemplateFile.CopyToAsync(filestream);
                    }

                    //Upload csv file to server
                    valuesFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + mailMergeFiles.ValuesFile.FileName;
                    using (var filestream = new FileStream(Path.Combine(path, valuesFileName), FileMode.Create))
                    {
                        await mailMergeFiles.ValuesFile.CopyToAsync(filestream);
                    }


                    //Read CSV file uploaded 
                    System.Data.DataTable dt = CSVReader.ConvertCSVtoDataTable(Path.Combine(path, valuesFileName));
                    wordMailMerge.generateEmail(Path.Combine(path, templateFileName), dt, Path.Combine(path, valuesFileName));
                    TempData["msg"] = "Email Generated successfully.";
                }
                else
                {
                    TempData["msg"] = "File doesnot contain any data.";
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally {
                try
                {
                    System.IO.File.Delete(Path.Combine(path, templateFileName));
                    System.IO.File.Delete(Path.Combine(path, valuesFileName));
                }
                catch(Exception)
                { }
            }
            return true;
        }
    }
}

