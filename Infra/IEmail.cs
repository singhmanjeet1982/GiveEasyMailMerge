using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Giveaway.Infra
{
    public interface IEmailSendGrid
    {
        public Task SendEmailAsync(string email, string subject, string message);
        public Task Execute(string email, string subject, string message);
        
    }
}
