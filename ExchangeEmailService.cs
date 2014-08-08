// --------------------------------------------------------------------------------------------------------------------
// <summary>
//   Provides access to Microsoft Exchange.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExchangeEmailProvider
{
    using System;
    using System.Net;
    
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Provides access to Microsoft Exchange.
    /// </summary>
    public static class ExchangeEmailService
    {
        /// <summary>
        /// Initializes static members of the <see cref="ExchangeEmailService"/> class.
        /// </summary>
        static ExchangeEmailService()
        {
            CertificateValidator.Initialize();
        }

        /// <summary>
        /// Returns a connected <see cref="ExchangeService"/>.
        /// </summary>
        /// <param name="user">The user.</param>
        /// <param name="listener">The trace listener.</param>
        /// <returns>An <see cref="ExchangeService"/> which is connected using impersonation.</returns>
        public static ExchangeService ConnectToService(ExchangeUser user, ITraceListener listener = null)
        {
            var service = new ExchangeService(user.Version);

            if (listener != null)
            {
                service.TraceListener = listener;
                service.TraceFlags = TraceFlags.All;
                service.TraceEnabled = true;
            }

            service.Credentials = new NetworkCredential(user.EmailAddress, user.Password);

            if (user.AutodiscoverUrl == null)
            {
                Console.Write("Using Autodiscover to find EWS URL for {0}. Please wait... ", user.EmailAddress);

                service.AutodiscoverUrl(user.EmailAddress, RedirectionUrlValidationCallback);
                user.AutodiscoverUrl = service.Url;

                Console.WriteLine("Autodiscover Complete");
            }
            else
            {
                service.Url = user.AutodiscoverUrl;
            }

            Console.WriteLine("Service Address: {0}", service.Url);
            return service;
        }

        /// <summary>
        /// Returns an <see cref="ExchangeService"/> which is connected using impersonation.
        /// </summary>
        /// <param name="user">The user.</param>
        /// <param name="impersonatedUserSmtpAddress">The impersonation address.</param>
        /// <param name="listener">The trace listener.</param>
        /// <returns>An <see cref="ExchangeService"/> which is connected using impersonation.</returns>
        public static ExchangeService ConnectToServiceWithImpersonation(ExchangeUser user, string impersonatedUserSmtpAddress, ITraceListener listener = null)
        {
            var service = new ExchangeService(user.Version);

            if (listener != null)
            {
                service.TraceListener = listener;
                service.TraceFlags = TraceFlags.All;
                service.TraceEnabled = true;
            }

            service.Credentials = new NetworkCredential(user.EmailAddress, user.Password);

            var impersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonatedUserSmtpAddress);

            service.ImpersonatedUserId = impersonatedUserId;

            if (user.AutodiscoverUrl == null)
            {
                service.AutodiscoverUrl(user.EmailAddress, RedirectionUrlValidationCallback);
                user.AutodiscoverUrl = service.Url;
            }
            else
            {
                service.Url = user.AutodiscoverUrl;
            }

            Console.WriteLine("Service Address: {0}", service.Url);
            return service;
        }

        /// <summary>
        /// Returns a value indicating whether or not the provided <paramref name="redirectionUrl"/> is valid.
        /// </summary>
        /// <param name="redirectionUrl">The redirection url.</param>
        /// <returns>
        /// A value indicating whether or not the provided <paramref name="redirectionUrl"/> is valid.
        /// </returns>
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return new Uri(redirectionUrl).Scheme == "https";
        }
    }
}
