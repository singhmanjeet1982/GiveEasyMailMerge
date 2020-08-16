using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Giveaway.Infra
{
    public interface IWordMailMerge
    {
        public string generateEmail(string templateFilePath, System.Data.DataTable templateValues, string valuesFilePath);
    }
}
