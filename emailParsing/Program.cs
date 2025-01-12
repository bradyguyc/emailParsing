using Azure.Identity;

using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text.RegularExpressions;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Concurrent;

namespace EmailParser
{
    internal class Program
    {
        private static string RemoveComments(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            return Regex.Replace(input, "<!--.*?-->", string.Empty, RegexOptions.Singleline);
        }

        private static string StripHtml(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            // Remove HTML tags using a regular expression
            return Regex.Replace(input, "<.*?>", string.Empty);
        }

        private static async Task ProcessMessageAsync(Message message, string emailDirectory)
        {
            var subject = message.Subject ?? "(No Subject)";
            var bodyContent = RemoveComments(StripHtml(message.Body?.Content ?? string.Empty));
            var toEmails = string.Join(", ", message?.ToRecipients?.Select(r => r.EmailAddress?.Address) ?? Enumerable.Empty<string>());
            var fromEmails = message?.From?.EmailAddress?.Address ?? string.Empty;

            var fileName = Path.Combine(emailDirectory, $"{message.Id}.txt");
            var emailContent = $"to:{toEmails}\nfrom:{fromEmails}\n\nSubject: {subject}\n\n{bodyContent}";

            await File.WriteAllTextAsync(fileName, emailContent);
        }

        private static async Task ProcessMessagePageAsync(IList<Message> messages, string emailDirectory, int pageNumber)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            foreach (var item in messages)
            {
                ProcessMessageAsync(item, emailDirectory);
                
            } 
           
            var elapsedTime = stopwatch.ElapsedMilliseconds;
            Console.WriteLine($"Page {pageNumber}: Time taken to process messages: {elapsedTime} ms");
            Debug.WriteLine($"Page {pageNumber}: Time taken to process messages: {elapsedTime} ms");
        }

        private static async Task Main(string[] args)
        {
            // Replace with your actual client ID
            var clientId = "83658fa9-7bf1-4190-ac7e-45d02cf514f5";

            // Directory to save emails
            var emailDirectory = @"c:\data\emails";

            // Date range for filtering emails
            var startDate = new DateTime(2024, 10, 1);
            var endDate = new DateTime(2025, 1, 31);

            // Initialize the interactive browser credential
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = "common",
                ClientId = clientId,
                RedirectUri = new Uri("http://localhost")
            };
            var credential = new InteractiveBrowserCredential(options);

            // Initialize Graph client with the credential and scopes
            var scopes = new[] { "Mail.Read" };
            var graphClient = new GraphServiceClient(credential, scopes);

            try
            {
                // Ensure the email directory exists
                if (!Directory.Exists(emailDirectory))
                {
                    Directory.CreateDirectory(emailDirectory);
                }

                // Format dates to match expected format
                var startDateString = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");
                var endDateString = endDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

             

                // Fetch the first page of messages
                var messagePage = await graphClient.Me.Messages
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Filter = $"receivedDateTime ge {startDateString} and receivedDateTime le {endDateString}";
                        requestConfiguration.QueryParameters.Select = new[] { "subject", "body", "toRecipients", "from" };
                        requestConfiguration.QueryParameters.Top = 100; // Specify the number of messages per page
                    });

                // Process messages and handle pagination
                int pageNumber = 1;

                while (messagePage != null)
                {
                    if (messagePage.Value != null)
                    {
                        var currentPage = messagePage;
                        var currentPageNumber = pageNumber;
                        Task.Run(async () => await ProcessMessagePageAsync(currentPage.Value, emailDirectory, currentPageNumber));
                        pageNumber++;
                    }

                    // Get the next page of messages, if any
                    if (!string.IsNullOrEmpty(messagePage.OdataNextLink))
                    {
                        // Use the OdataNextLink to fetch the next page
                        messagePage = await graphClient
                            .Me
                            .Messages
                            .WithUrl(messagePage.OdataNextLink)
                            .GetAsync();
                    }
                    else
                    {
                        messagePage = null;
                    }
                }

              

                Console.WriteLine("Emails have been saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
