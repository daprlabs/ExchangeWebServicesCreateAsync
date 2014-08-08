// --------------------------------------------------------------------------------------------------------------------
// <summary>
//   The certificate validator.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExchangeEmailProvider
{
    using System.Linq;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;

    /// <summary>
    /// The certificate validator.
    /// </summary>
    public static class CertificateValidator
    {
        /// <summary>
        /// Initializes this class.
        /// </summary>
        public static void Initialize()
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
        }

        /// <summary>
        /// Validates the provided certificate.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="certificate">
        /// The certificate.
        /// </param>
        /// <param name="chain">
        /// The chain.
        /// </param>
        /// <param name="sslPolicyErrors">
        /// The SSL policy errors.
        /// </param>
        /// <returns>
        /// A value indicating whether or not the provided certificate is valid.
        /// </returns>
        private static bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == 0)
            {
                return false;
            }

            if (chain != null)
            {
                return
                    chain.ChainStatus.Where(status => (certificate.Subject != certificate.Issuer) || (status.Status != X509ChainStatusFlags.UntrustedRoot))
                        .All(status => status.Status == X509ChainStatusFlags.NoError);
            }

            return true;
        }
    }
}
