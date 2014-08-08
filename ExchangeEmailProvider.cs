// --------------------------------------------------------------------------------------------------------------------
// <summary>
//   Defines the ExchangeEmailProvider type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExchangeEmailProvider
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security;

    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Methods for sending email using Microsft Exchange.
    /// </summary>
    public class ExchangeEmailProvider
    {
        private static void Main(string[] args)
        {
            var password = new SecureString();
            (Console.ReadLine() ?? string.Empty).ToList().ForEach(password.AppendChar);
            var user = new ExchangeUser { EmailAddress = "from@address.com", Password = password, Version = ExchangeVersion.Exchange2013 };
            var service = ExchangeEmailService.ConnectToService(user);

            // Create an email and add some recipients
            var message = new EmailMessage(service) { Subject = "GOTTA GO FAST", Body = "Check it out! https://www.youtube.com/watch?v=Qzb87qir7Cs", Importance = Importance.High };
            message.ToRecipients.Add("AAAAAAAAAA@CCCCCCCCCCCCCC.TLD");
            message.ToRecipients.Add("BBBBBBBBBB@BBBBBBBBBBBBBB.TLD");
            message.ToRecipients.Add("CCCCCCCCCC@AAAAAAAAAAAAAA.TLD");

            try
            {
                // Asynchronously send the email... but block on the result because this is just a test..
                var responses = service.CreateItemsAsync(new List<EmailMessage> { message }, WellKnownFolderName.Drafts, MessageDisposition.SendOnly, null).Result;

                // Check the response to determine whether the email messages were successfully submitted.
                if (responses.OverallResult == ServiceResult.Success)
                {
                    Console.WriteLine("All email messages were successfully submitted");
                    return;
                }

                foreach (var response in responses)
                {
                    Console.WriteLine(
                        "Result: {0}\nError Code: {1}\nError Message: {2}",
                        response.Result,
                        response.ErrorCode,
                        response.ErrorMessage);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("Press or select Enter...");
            Console.ReadLine();
        }
    }
}
