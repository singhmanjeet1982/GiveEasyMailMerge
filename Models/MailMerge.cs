using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using Giveaway.Helper;

namespace Giveaway.Models
{
    public class MailMerge
    {
        [Display(Name = "Template File:")]
        [FileExtension("dotx", ErrorMessage = "Please select only word template (.dotx) file.")]
        public IFormFile TemplateFile { get; set; }

        [Display(Name = "CSV File:")]
        [FileExtension("csv", ErrorMessage = "Please select only csv (.csv) file.")]
        public IFormFile ValuesFile { get; set; }
    }
}
