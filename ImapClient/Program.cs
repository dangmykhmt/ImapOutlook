using Spire.Email;
using Spire.Email.IMap;
using Spire.Email.Pop3;
using Spire.Email.Smtp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImapClient
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Create a POP3 client  
            Pop3Client pop = new Pop3Client();
            //Set host, username, password etc. for the client  
            pop.Host = "outlook.office365.com";
            pop.Username = "buttsleggs7@hotmail.com";
            pop.Password = "SVU9dP24";
            pop.Port = 995;
            pop.EnableSsl = true;
            //Connect the server  
            pop.Connect();
            //Get the first message by its sequence number  
            MailMessage message = pop.GetMessage(2);
            var messageCount = pop.GetMessageCount();
            //Parse the message  
            Console.WriteLine("------------------ HEADERS ---------------");
            Console.WriteLine("From : " + message.From.ToString());
            Console.WriteLine("To : " + message.To.ToString());
            Console.WriteLine("Date : " + message.Date.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Subject: " + message.Subject);
            Console.WriteLine("------------------- BODY -----------------");
            Console.WriteLine(message.BodyText);
            Console.WriteLine("------------------- END ------------------");
            //Save the message to disk using its subject as file name  
            message.Save(message.Subject + ".eml", MailMessageFormat.Eml);
            Console.WriteLine("Message Saved.");
        }
    }
}
