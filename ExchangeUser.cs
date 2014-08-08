// --------------------------------------------------------------------------------------------------------------------
// <summary>
//   Represents a user in Exchange.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExchangeEmailProvider
{
    using System;
    using System.Security;

    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a user in Exchange.
    /// </summary>
    public class ExchangeUser
    {
        /// <summary>
        /// Gets or sets the version.
        /// </summary>
        public ExchangeVersion Version { get; set; }

        /// <summary>
        /// Gets or sets the email address.
        /// </summary>
        public string EmailAddress { get; set; }

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
        public SecureString Password { get; set; }

        /// <summary>
        /// Gets or sets the autodiscover url.
        /// </summary>
        public Uri AutodiscoverUrl { get; set; }
    }
}