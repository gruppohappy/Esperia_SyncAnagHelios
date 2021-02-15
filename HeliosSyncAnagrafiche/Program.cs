using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace HeliosSyncAnagrafiche
{
    class Program
    {
        public string connectionString = @"Data Source=SRVSRDESP\SQLSRDESP; Initial Catalog=Helios_Esperia;User=sa;Password=esp2015*";
        public string logFolderName = "SyncAnagraficheLogs";

        static void Main(string[] args)
        {
            // Instazio l'oggetto di classe syncService
            SyncService syncService = new SyncService();
        }

        public void syncUDCCore()
        {
            SqlConnection conn = null;

            try
            {
                conn = new SqlConnection(connectionString);
                conn.Open();
                //WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Connessione con SrvSrdEsperia aperta correttamente"));

                try
                {
                    string Query = string.Concat(" INSERT INTO [Helios_Esperia].[dbo].[AnagraficaUdc] ",
                                                  "SELECT 'U'+a.[UDC],a.[ARTICOLO],a.[UM],a.[QUALITA] ",
                                                  "from [SRVTSESP\\SQLESPERIA].[ESPERIA_SL].[dbo].[vSL_Giacenza] a LEFT JOIN ",
                                                  "AnagraficaUdc b ON a.udc=SUBSTRING(b.UDC,2,18) ",
                                                  "WHERE isnull(b.udc,'')='' AND len(a.[UDC])=18");
                    SqlCommand cmd = new SqlCommand(Query, conn);
                    int result = cmd.ExecuteNonQuery();

                    if (result > 0)
                    {
                        WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - ", result.ToString(), " righe inserite correttamente"));
                    }

                }
                catch (Exception ex)
                {
                    WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Errore in fase di sync, invio mail. Dettagli", ":", ex.ToString()));
                    sendmail("fabio.palmisano@gruppo-happy.it", "SYNC UDC SRVSRD ESPERIA", string.Concat("ERRORE IN FASE DI SYNC UDC SU SRVSRDESP:", ex.ToString()));
                }

            }
            catch (Exception ex)
            {
                WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Errore in fase di sync (step sync core), invio mail. Dettagli", ":", ex.ToString()));
                sendmail("fabio.palmisano", "SYNC UDC SRVSRD ESPERIA", string.Concat("ERRORE IN FASE DI SYNC UDC SU SRVSRDESP:", ex.ToString()));
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    //WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Connessione con SrvSrdEsperia chiusa correttamente"));
                }
            }
        }

        public void syncAnagCore(string _connectionString)
        {
            //Core del servizio: sincronizza tutte le anagrafiche in un colpo solo e restituisce gli esiti nel log.
            SqlConnection conn = null;

            try
            {
                conn = new SqlConnection(_connectionString);
                conn.Open();

                //WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Connessione con SrvSrdEsperia aperta correttamente"));

                syncAnagrafica("FOGLIE", ref conn);
                syncAnagrafica("BANCALI", ref conn);
                syncAnagrafica("CALZE", ref conn);
                syncAnagrafica("CASSE", ref conn);
                syncAnagrafica("SACCHI", ref conn);
                syncAnagrafica("SCATOLE", ref conn);
                syncAnagrafica("ASSORBENTI", ref conn);


            }
            catch (Exception ex)
            {
                WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - (SYNC ANAGRAFICHE) Errore in fase di sync anagrafica (step sync core)", ":", ex.ToString()));
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    //WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Connessione con SrvSrdEsperia chiusa correttamente"));
                }
            }
        }

        public void syncAnagrafica(string _syncStep, ref SqlConnection _conn)
        {
            try
            {
                string Query = string.Concat(" INSERT INTO [Helios_Esperia].[dbo].[AnagraficaMateriali]",
                                             " SELECT  ANAG_AS400.* FROM [dbo].[vAs400ANAG_", _syncStep, "] ANAG_AS400",
                                             " LEFT JOIN [Helios_Esperia].[dbo].[AnagraficaMateriali] ANAG_HELIOS ",
                                             " ON ANAG_AS400.CODICE=ANAG_HELIOS.CodiceMateriale ",
                                             " WHERE ISNULL(ANAG_HELIOS.CodiceMateriale,'')='' ");
                SqlCommand cmd = new SqlCommand(Query, _conn);
                int result = cmd.ExecuteNonQuery();

                if (result > 0)
                {
                    WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Anagrafica ", _syncStep, " ", result.ToString(), " righe inserite correttamente"));
                }
            }
            catch (Exception ex)
            {
                sendmail("fabio.palmisano@gruppo-happy.it", "SYNC ANAGRAFICHE SRVSRD ESPERIA", string.Concat("ERRORE IN FASE DI SYNC ANAGRAFICHE SU SRVSRDESP:", ex.ToString()));
                WriteToFile(logFolderName, string.Concat(DateTime.Now.ToString(), " - Errore in fase di sync anagrafica ", _syncStep, ", invio mail. Dettagli: ", ex.ToString()));
            }
        }

        public bool connectionCheck(string connectionString)
        {
            SqlConnection conn = null;

            try
            {

                conn = new SqlConnection(connectionString);
                conn.Open();
                return true;
            }
            catch (Exception ex)
            {
                WriteToFile("SeviceLogs", string.Concat(string.Concat(DateTime.Now.ToString(), " - Errore di connessione: ", ex.ToString())));
                sendmail("fabio.palmisano@gruppo-happy.it", "CONNECTION CHECK", string.Concat("ERRORE DI CONNESSIONE CON ", connectionString, ": ", ex.ToString()));
                return false;
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        public void WriteToFile(string Folder, string Message)
        {
            string path = string.Concat(AppDomain.CurrentDomain.BaseDirectory, "\\", Folder);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = string.Concat(AppDomain.CurrentDomain.BaseDirectory, "\\", Folder, "\\ServiceLog_", DateTime.Now.Date.ToShortDateString().Replace('/', '_'), ".txt");
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        public bool sendmail(string recievermailid, string subject, string bodyText)
        {
            try
            {
                string senderId = "support@gruppo-happy.it"; // Sender EmailID
                string senderPassword = "7t9Pe!aB"; // Sender Password

                System.Net.Mail.MailMessage mailMessage = new System.Net.Mail.MailMessage();
                mailMessage.To.Add(recievermailid);
                mailMessage.From = new MailAddress(senderId);

                mailMessage.Subject = subject;
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;

                mailMessage.Body = bodyText;
                mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
                mailMessage.IsBodyHtml = false;

                mailMessage.Priority = MailPriority.High;

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Credentials = new System.Net.NetworkCredential(senderId, senderPassword);
                smtpClient.Port = 587;
                smtpClient.Host = "smtp.gmail.com";
                smtpClient.EnableSsl = true;

                object userState = mailMessage;

                try
                {
                    smtpClient.Send(mailMessage);
                    return true;
                }
                catch (System.Net.Mail.SmtpException)
                {
                    WriteToFile("SeviceLog", string.Concat(DateTime.Now.ToString(), " - Errore in fase di invio mail, SMTP EXCEPTION"));
                    return false;
                }
            }
            catch (Exception ex)
            {
                WriteToFile("SeviceLog", string.Concat(DateTime.Now.ToString(), " - Errore in fase di invio mail, dettagli:", ex.ToString()));
                return false;
            }
        }
    }

    internal class SyncService : Program
    {
        public SyncService()
        {
            // SyncAnagraficheHelios
            try
            {
                if (connectionCheck(connectionString))
                {
                    syncAnagCore(connectionString);
                }
                else
                {
                    WriteToFile(logFolderName, string.Concat(DateTime.Now, " Sevizio avviato ma il sync anagrafica non verrà eseguito. Problemi lato connessione. ",
                                                           "Attendere il prossimo ciclo."));
                }

                // SyncAnagraficaUDC
                if (connectionCheck(connectionString))
                {
                    syncUDCCore();
                }
                else
                {
                    WriteToFile(logFolderName, string.Concat(DateTime.Now, " - Sevizio avviato ma il sync Udc non verrà eseguito. Problemi lato connessione. ",
                                                                           "Attendere il prossimo ciclo."));
                }
            }
            catch (Exception ex)
            {
                WriteToFile(logFolderName, string.Concat(DateTime.Now, " Errore: ", ex.ToString()));
            }
        }


    }
}
