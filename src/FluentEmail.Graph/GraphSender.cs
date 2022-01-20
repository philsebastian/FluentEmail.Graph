namespace FluentEmail.Graph
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Azure.Identity;
    using FluentEmail.Core;
    using FluentEmail.Core.Interfaces;
    using FluentEmail.Core.Models;
    using JetBrains.Annotations;
    using Microsoft.Graph;

    /// <summary>
    /// Implementation of <c>ISender</c> for the Microsoft Graph API.
    /// See <see cref="FluentEmailServicesBuilderExtensions"/>.
    /// </summary>
    public class GraphSender : ISender
    {
        private readonly bool saveSent;

        private readonly GraphServiceClient graphClient;

        public GraphSender(GraphSenderOptions options)
        {
            this.saveSent = options.SaveSentItems ?? true;

            ClientSecretCredential spn = new ClientSecretCredential(options.TenantId, options.ClientId, options.Secret);

            this.graphClient = new(spn);
        }

        public SendResponse Send(IFluentEmail email, CancellationToken? token = null)
        {
            return this.SendAsync(email, token).GetAwaiter().GetResult();
        }

        public async Task<SendResponse> SendAsync(IFluentEmail email, CancellationToken? token = null)
        {
            int minimumSizeLimit = 1024 * 1024 * 3;
            try
            {
                var rawMessage = CreateMessage(email);

                var message = await this.graphClient.Users[email.Data.FromAddress.EmailAddress].Messages.Request().AddAsync(rawMessage);

                if (email.Data.Attachments != null && email.Data.Attachments.Count > 0)
                {
                    foreach (var attachment in email.Data.Attachments)
                    {
                        await UploadFileAttachment(attachment);
                    }
                }

                await this.graphClient.Users[email.Data.FromAddress.EmailAddress].Messages[message.Id].Send().Request().PostAsync();

                return new SendResponse
                {
                    MessageId = message.Id,
                };

                async Task UploadFileAttachment(Core.Models.Attachment a)
                {
                    var theBytes = GetAttachmentBytes(a.Data);
                    if (theBytes.Length < minimumSizeLimit)
                    {
                        await UploadSmall();
                    }
                    else
                    {
                        await UploadLarge();
                    }

                    async Task UploadLarge()
                    {
                        var attachment = new AttachmentItem
                        {
                            Name = a.Filename,
                            AttachmentType = AttachmentType.File,
                            Size = theBytes.Length
                        };

                        var uploadSession = await this.graphClient.Users[email.Data.FromAddress.EmailAddress].Messages[message.Id].Attachments.CreateUploadSession(attachment).Request().PostAsync();

                        var largeFileUpload = new LargeFileUploadTask<AttachmentItem>(uploadSession, a.Data);

                        var uploadedFile = await largeFileUpload.UploadAsync();
                        var success = uploadedFile.UploadSucceeded;
                    }

                    async Task UploadSmall()
                    {
                        var attachment = new FileAttachment
                        {
                            Name = a.Filename,
                            ContentType = a.ContentType,
                            ContentBytes = theBytes
                        };

                        await this.graphClient.Users[email.Data.FromAddress.EmailAddress].Messages[message.Id].Attachments.Request().AddAsync(attachment);
                    }
                }
            }
            catch (Exception ex)
            {
                return new SendResponse
                {
                    ErrorMessages = new List<string> { ex.Message },
                };
            }
        }
        private static Message CreateMessage(IFluentEmail email)
        {
            var messageBody = new ItemBody
            {
                Content = email.Data.Body,
                ContentType = email.Data.IsHtml ? BodyType.Html : BodyType.Text,
            };

            var message = new Message();
            message.Subject = email.Data.Subject;
            message.Body = messageBody;
            message.From = ConvertToRecipient(email.Data.FromAddress);
            message.ReplyTo = CreateRecipientList(email.Data.ReplyToAddresses);
            message.ToRecipients = CreateRecipientList(email.Data.ToAddresses);
            message.CcRecipients = CreateRecipientList(email.Data.CcAddresses);
            message.BccRecipients = CreateRecipientList(email.Data.BccAddresses);

            switch (email.Data.Priority)
            {
                case Priority.High:
                    message.Importance = Importance.High;
                    break;

                case Priority.Normal:
                    message.Importance = Importance.Normal;
                    break;

                case Priority.Low:
                    message.Importance = Importance.Low;
                    break;

                default:
                    message.Importance = Importance.Normal;
                    break;
            }

            return message;
        }

        private static IList<Recipient> CreateRecipientList(IEnumerable<Address> addressList)
        {
            if (addressList == null)
            {
                return new List<Recipient>();
            }

            return addressList
                .Select(ConvertToRecipient)
                .ToList();
        }

        private static Recipient ConvertToRecipient([NotNull] Address address)
        {
            if (address is null)
            {
                throw new ArgumentNullException(nameof(address));
            }

            return new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = address.EmailAddress,
                    Name = address.Name,
                },
            };
        }

        private static byte[] GetAttachmentBytes(Stream stream)
        {
            using var m = new MemoryStream();
            stream.CopyTo(m);
            return m.ToArray();
        }
    }
}
