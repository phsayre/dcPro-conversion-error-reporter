//#define MYDEBUG

using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace conversionErrorReporter
{
    /// <summary>
    /// Reports error to the database and sends out an email containing error details.
    /// </summary>
    class ErrorReporter
    {
        private string connString; 
        private string emailHost; 
        private string errorDirPath; 
        private string fromAddress;  //for SendEmail()
        private string toAddress;  //for SendEmail()
        // All of the above variables are specified in app.config
        private string itemID;
        private static NpgsqlConnection conn;
        private FileInfo[] errorItems;
        private DirectoryInfo[] errorItemFolders;


        public ErrorReporter()
        {
            Setup();
        }

        /// <summary>
        /// Assigns variables according to app.config settings.
        /// </summary>
        private void Setup()
        {
            try
            {
                connString = ConfigurationManager.AppSettings["connString"];
                emailHost = ConfigurationManager.AppSettings["emailHost"];
                errorDirPath = ConfigurationManager.AppSettings["errorDirPath"];
                fromAddress = ConfigurationManager.AppSettings["fromAddress"];
                toAddress = ConfigurationManager.AppSettings["toAddress"];
                DirectoryInfo errorDir = new DirectoryInfo(errorDirPath);
                errorItemFolders = errorDir.GetDirectories("*.*");
                errorItems = errorDir.GetFiles("*");
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("There was a problem accessing the directory or retrieving the app settings; check app.config: " + e.Message);                
            }
        }

        /// <summary>
        /// Checks if the error directory is empty.
        /// </summary>
        /// <returns>True when empty.</returns>
        private bool IsEmpty()
        {
            bool isEmptyFiles = !Directory.EnumerateFiles(errorDirPath).Any();
            bool isEmptyFolders = !Directory.EnumerateDirectories(errorDirPath).Any();
            return (isEmptyFiles && isEmptyFolders);
        }

        /// <summary>
        /// Checks if the file has already been reported.
        /// </summary>
        /// <param name="file">The file to check.</param>
        /// <returns>True if file has not been reported to database.</returns>
        private bool IsReported(FileInfo file)
        {
            bool reported;
            try
            {
                itemID = (Path.GetFileNameWithoutExtension(file.FullName));
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                com.CommandText = String.Format("SELECT converter_errormsg FROM redacted.items WHERE id={0};", itemID);
                string pgResponse = com.ExecuteScalar().ToString();
                if (string.IsNullOrWhiteSpace(pgResponse))
                {
                    reported = false;
                }
                else
                {
                    reported = true;
                }
                conn.Close();
                return reported;
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Error connecting to the database: " + e.Message);
                return true;
            }
        }

        /// <summary>
        /// Checks if the folder has already been reported.
        /// </summary>
        /// <param name="folder">The folder to check.</param>
        /// <returns>True if the folder has not been reported to the database.</returns>
        //overloaded to handle folder
        private bool IsReported(DirectoryInfo folder)
        {
            bool reported;
            try
            {
                itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                com.CommandText = String.Format("SELECT converter_errormsg FROM redacted.items WHERE id={0};", itemID);
                string pgResponse = com.ExecuteScalar().ToString();
                if (string.IsNullOrWhiteSpace(pgResponse))
                {
                    reported = false;
                }
                else
                {
                    reported = true;
                }
                conn.Close();
                return reported;
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Error connecting to the database: " + e.Message);
                return true;
            }
        }


        private static bool SpecialBackslashes()
        {
            IDbCommand com = conn.CreateCommand();
            com.CommandText = "show standard_conforming_strings;";
            string standard_strings = com.ExecuteScalar().ToString();

            return (standard_strings == "off");
        }


        private static string SqlEscape(string query)
        {
            query = query.Replace("'", "''");
            if (SpecialBackslashes())
                query = query.Replace(@"\", @"\\");
            return query;
        }

        /// <summary>
        /// Generates the error message for the email.
        /// </summary>
        /// <returns>A string message.</returns>
        private string GenerateErrorMessage(List<FileInfo> errorItemsToReport, List<DirectoryInfo> errorItemFoldersToReport)
        {
            string emailBody = "Errors occurred while DC Pro attempted to convert the following\r";

            if (errorItemsToReport.Count != 0)
            {
                emailBody += "\r\nfiles:\r";
                foreach (FileInfo file in errorItemsToReport)
                {
                    emailBody += "\r\n";
                    emailBody += errorItemsToReport.IndexOf(file).ToString() + ") ";
                    emailBody += file.FullName;
                    
                    try
                    {
                        itemID = (Path.GetFileNameWithoutExtension(file.FullName));
                        conn = new NpgsqlConnection(connString);
                        conn.Open();
                        NpgsqlCommand com = conn.CreateCommand();
                        com.CommandText = String.Format("SELECT converter_errormsg FROM redacted.items WHERE id={0};", itemID);
                        string pgResponse = com.ExecuteScalar().ToString();
                        conn.Close();
                        emailBody += " | " + pgResponse;
                    }
                    catch
                    {
                        //do nothing
                    }
                }
            }

            if (errorItemFoldersToReport.Count != 0)
            {
                emailBody += "\r\n\r\nfolders:\r";
                foreach (DirectoryInfo folder in errorItemFoldersToReport)
                {
                    emailBody += "\r\n";
                    emailBody += errorItemFoldersToReport.IndexOf(folder).ToString() + ") ";
                    emailBody += folder.FullName;

                    try
                    {
                        itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
                        conn = new NpgsqlConnection(connString);
                        conn.Open();
                        NpgsqlCommand com = conn.CreateCommand();
                        com.CommandText = String.Format("SELECT converter_errormsg FROM redacted.items WHERE id={0};", itemID);
                        string pgResponse = com.ExecuteScalar().ToString();
                        conn.Close();
                        emailBody += " | " + pgResponse;
                    }
                    catch
                    {
                        //do nothing
                    }
                }
            }

            return emailBody;
        }

        /// <summary>
        /// Sends the email using smtp.
        /// </summary>
        /// <param name="emailBody">The message to send.</param>
        private void SendEmail(string emailBody)
        {
            string emailSubject = "PDF Neevia Converter Error: " + DateTime.Now.ToString();

            try
            {
                SmtpClient myClient = new SmtpClient(emailHost);
                MailMessage myMessage = new MailMessage(fromAddress, toAddress, emailSubject, emailBody);
                myClient.Send(myMessage);
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Error while attempting to send mail: " + e.Message);
            }
        }

        /// <summary>
        /// (Files) Updates the database and records the file converter_error=true.
        /// </summary>
        /// <param name="file">File in error.</param>
        private void UpdateDb(FileInfo file)
        {
            itemID = (Path.GetFileNameWithoutExtension(file.FullName));

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                string msg = "Failed to convert:" + " (" + errorDirPath + file.Name + ")";
                com.CommandText = String.Format("UPDATE redacted.items SET converted=false, converter_error=true, converter_errormsg='{1}' WHERE id={0};", itemID, SqlEscape(msg));
                com.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
            }
        }

        /// <summary>
        /// (Folders) Updates the database and records the folder error_converter=true.
        /// </summary>
        /// <param name="folder">Folder in error.</param>
        //overloaded to handle a folder, instead of a file
        private void UpdateDb(DirectoryInfo folder)
        {
            itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                string msg = "Failed to convert:" + " (" + errorDirPath + folder.Name + ")";
                com.CommandText = String.Format("UPDATE redacted.items SET converted=false, converter_error=true, converter_errormsg='{1}' WHERE id={0};", itemID, SqlEscape(msg));
                com.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
            }
        }

        /// <summary>
        /// Runs the program in logical order.
        /// Does nothing when the error directory is empty.
        /// Takes files and then folders one at a time and updates the database.
        /// Sends an email out with the error details.
        /// </summary>
        public void Run()
        {
            if (IsEmpty())
            {
                //WriteOut.HandleMessage("Error folder is empty. Nothing to report.");
                //do nothing
            }
            else
            {
                List<FileInfo> errorItemsToReport = new List<FileInfo>();
                List<DirectoryInfo> errorItemFoldersToReport = new List<DirectoryInfo>();

                foreach (FileInfo file in errorItems)
                {
                    if (!IsReported(file))
                    {
                        UpdateDb(file);
                        errorItemsToReport.Add(file);
                    }
                }
                foreach (DirectoryInfo folder in errorItemFolders)
                {
                    if (!IsReported(folder))
                    {
                        UpdateDb(folder);
                        errorItemFoldersToReport.Add(folder);
                    }
                }

                if (errorItemsToReport.Count != 0 || errorItemFoldersToReport.Count != 0)
                {
                    WriteOut.HandleMessage(errorItems.Length.ToString() + " file(s) and " + errorItemFolders.Length.ToString() + " folder(s) found in the error folder.");
                    SendEmail(GenerateErrorMessage(errorItemsToReport, errorItemFoldersToReport));
                }
#if MYDEBUG
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
#endif
            }         
        }
    }
}
