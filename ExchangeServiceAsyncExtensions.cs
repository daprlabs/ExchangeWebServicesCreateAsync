// --------------------------------------------------------------------------------------------------------------------
// <summary>
//   Asynchronous extension methods for <see cref="ExchangeService" />.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExchangeEmailProvider
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Asynchronous extension methods for <see cref="ExchangeService"/>.
    /// </summary>
    public static class ExchangeServiceAsyncExtensions
    {
        /// <summary>
        /// The Asynchronous Programming Model begin-operation method for CreateItems.
        /// </summary>
        private static readonly Lazy<Func<ExchangeService, IEnumerable<Item>, FolderId, MessageDisposition?, SendInvitationsMode?, IAsyncResult>> BeginCreateItems =
            new Lazy<Func<ExchangeService, IEnumerable<Item>, FolderId, MessageDisposition?, SendInvitationsMode?, IAsyncResult>>(GenerateBeginCreateItems);

        /// <summary>
        /// The Asynchronous Programming Model end-operation method for CreateItems.
        /// </summary>
        private static readonly Lazy<Func<ExchangeService, IAsyncResult, ServiceResponseCollection<ServiceResponse>>> EndCreateItems = new Lazy<Func<ExchangeService, IAsyncResult, ServiceResponseCollection<ServiceResponse>>>(GenerateEndCreateItems);
        
        /// <summary>
        /// Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact, PostItem, Task and Item.
        /// CreateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="service">
        /// The service instance.
        /// </param>
        /// <param name="items">
        /// The items to create.
        /// </param>
        /// <param name="parentFolderId">
        /// The Id of the folder in which to place the newly created items. If null, items are created in their default folders.
        /// </param>
        /// <param name="messageDisposition">
        /// Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.
        /// </param>
        /// <param name="sendInvitationsMode">
        /// Indicates if and how invitations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.
        /// </param>
        /// <returns>
        /// A ServiceResponseCollection providing creation results for each of the specified items.
        /// </returns>
        public static System.Threading.Tasks.Task<ServiceResponseCollection<ServiceResponse>> CreateItemsAsync(
            this ExchangeService service,
            IEnumerable<Item> items,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode)
        {
            return System.Threading.Tasks.Task.Factory.FromAsync(
                BeginCreateItems.Value(service, items, parentFolderId, messageDisposition, sendInvitationsMode),
                asyncResult => EndCreateItems.Value(service, asyncResult));
        }

        /// <summary>
        /// Returns the begin-operation method for the asynchronous counterpart to <see cref="ExchangeService.CreateItems"/>.
        /// </summary>
        /// <returns>The begin-operation method for the asynchronous counterpart to <see cref="ExchangeService.CreateItems"/>.</returns>
        private static Func<ExchangeService, IEnumerable<Item>, FolderId, MessageDisposition?, SendInvitationsMode?, IAsyncResult> GenerateBeginCreateItems()
        {
            // Parameters.
            var serviceParam = Expression.Parameter(typeof(ExchangeService), "service");
            var itemsParam = Expression.Parameter(typeof(IEnumerable<Item>), "items");
            var folderIdParam = Expression.Parameter(typeof(FolderId), "folderId");
            var messageDispositionParam = Expression.Parameter(typeof(MessageDisposition?), "messageDisposition");
            var sendInvitationsModeParam = Expression.Parameter(typeof(SendInvitationsMode?), "sendInvitationsMode");
            var createItemRequestType = typeof(ExchangeService).Assembly.GetType("Microsoft.Exchange.WebServices.Data.CreateItemRequest");
            var serviceErrorHandlingType = typeof(ExchangeService).Assembly.GetType("Microsoft.Exchange.WebServices.Data.ServiceErrorHandling");
            var createItemRequest = Expression.Parameter(createItemRequestType, "createItemRequest");
            var serviceErrorHandling = Expression.Parameter(serviceErrorHandlingType, "serviceErrorHandling");

            // Construct the request object.
            var itemConstructorInfo = createItemRequestType.GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance).FirstOrDefault();
            if (itemConstructorInfo == null)
            {
                throw new MissingMethodException("Cannot find the constructor for " + createItemRequestType + ".");
            }

            var newItem = Expression.New(
                itemConstructorInfo,
                serviceParam,
                Expression.Convert(Expression.Constant(1), serviceErrorHandlingType));
            var bodyExpressions = new List<Expression> { Expression.Assign(createItemRequest, newItem) };

            // Assign all properties in the request object.
            var properties = new Dictionary<string, ParameterExpression>
                             {
                                 { "ParentFolderId", folderIdParam },
                                 { "Items", itemsParam },
                                 { "MessageDisposition", messageDispositionParam },
                                 { "SendInvitationsMode", sendInvitationsModeParam }
                             };
            foreach (var prop in properties)
            {
                var info = createItemRequestType.GetProperty(prop.Key);
                if (info == null)
                {
                    throw new MissingMethodException("Could not find " + prop.Key + " property on " + createItemRequestType + ".");
                }

                bodyExpressions.Add(Expression.Call(createItemRequest, info.GetSetMethod(), new Expression[] { prop.Value }));
            }

            var beginExecute = createItemRequestType.GetMethod(
                "BeginExecute",
                BindingFlags.NonPublic | BindingFlags.FlattenHierarchy | BindingFlags.Instance);
            var call = Expression.Call(
                createItemRequest,
                beginExecute,
                Expression.Constant(null, typeof(AsyncCallback)),
                Expression.Constant(null, typeof(object)));

            bodyExpressions.Add(call);
            var body = Expression.Block(new[] { createItemRequest, serviceErrorHandling }, bodyExpressions);

            var lambda = Expression.Lambda<Func<ExchangeService, IEnumerable<Item>, FolderId, MessageDisposition?, SendInvitationsMode?, IAsyncResult>>(
                body,
                serviceParam,
                itemsParam,
                folderIdParam,
                messageDispositionParam,
                sendInvitationsModeParam);
            var compiled = lambda.Compile();
            return compiled;
        }

        /// <summary>
        /// Returns the end-operation method for the asynchronous counterpart to <see cref="ExchangeService.CreateItems"/>.
        /// </summary>
        /// <returns>The end-operation method for the asynchronous counterpart to <see cref="ExchangeService.CreateItems"/>.</returns>
        private static Func<ExchangeService, IAsyncResult, ServiceResponseCollection<ServiceResponse>> GenerateEndCreateItems()
        {
            var serviceParam = Expression.Parameter(typeof(ExchangeService), "service");
            var asyncResultParam = Expression.Parameter(typeof(IAsyncResult), "asyncResult");

            var asyncRequestResultType = typeof(ExchangeService).Assembly.GetType("Microsoft.Exchange.WebServices.Data.AsyncRequestResult");
            var createItemRequestType = typeof(ExchangeService).Assembly.GetType("Microsoft.Exchange.WebServices.Data.CreateItemRequest");
            var extractServiceRequest =
                asyncRequestResultType.GetMethod("ExtractServiceRequest", BindingFlags.Static | BindingFlags.Public)
                    .MakeGenericMethod(createItemRequestType);

            var createItemRequest = Expression.Call(extractServiceRequest, serviceParam, asyncResultParam);
            var endExecute = createItemRequestType.GetMethod("EndExecute", BindingFlags.NonPublic | BindingFlags.FlattenHierarchy | BindingFlags.Instance);
            var call = Expression.Call(createItemRequest, endExecute, new Expression[] { asyncResultParam });
            var lambda = Expression.Lambda<Func<ExchangeService, IAsyncResult, ServiceResponseCollection<ServiceResponse>>>(call, serviceParam, asyncResultParam);
            return lambda.Compile();
        }
    }
}