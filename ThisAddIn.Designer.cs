using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EventAdder
{
	public partial class ThisAddIn
	{
		// GLOBAL VARIABLE: Controls the "Pause" feature
		public static DateTime PauseUntil = DateTime.MinValue;

		private Outlook.NameSpace _nameSpace;
		private Outlook.MAPIFolder _inbox;
		private Outlook.Items _inboxItems;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			_nameSpace = this.Application.GetNamespace("MAPI");
			_inbox = _nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
			_inboxItems = _inbox.Items;

			// Monitor Incoming Mail
			_inboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(InboxItems_ItemAdd);
		}

		// --- MAIN LOGIC ---
		private async void InboxItems_ItemAdd(object item)
		{
			if (item == null) return;

			// 1. CHECK PAUSE: Is the plugin paused?
			if (DateTime.Now < PauseUntil)
			{
				return; // Stop here, do not forward.
			}

			// 2. CHECK RULES (Fast): Is it already in trash?
			if (IsLocatedInDeletedItems(item)) return;

			try
			{
				// 3. WAIT: Give Outlook Rules 1.5 seconds to move the email
				await Task.Delay(1500);

				// 4. CHECK RULES (Authoritative): Is it in trash now?
				if (IsLocatedInDeletedItems(item)) return;

				// 5. PROCESS
				if (item is Outlook.MailItem mail)
				{
					ProcessItem(mail, mail.Sender, mail.Subject, mail.Body, mail.SentOn);
				}
				else if (item is Outlook.MeetingItem meeting)
				{
					dynamic dMeeting = meeting;
					Outlook.AddressEntry sender = dMeeting.Sender;
					ProcessItem(meeting, sender, meeting.Subject, meeting.Body, meeting.SentOn);
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"Error processing item: {ex.Message}");
			}
		}

		// --- HELPER METHODS ---

		private void ProcessItem(object sourceItem, Outlook.AddressEntry sender, string subject, string body, DateTime receivedTime)
		{
			if (sender == null) return;

			string targetEmail = Properties.Settings.Default.TargetEmail;
			if (string.IsNullOrEmpty(targetEmail)) return;

			string senderAddress = GetSenderAddress(sender);

			if (string.Equals(senderAddress, targetEmail, StringComparison.OrdinalIgnoreCase))
			{
				return;
			}

			CreateAndSendEvent(sourceItem, subject, body, receivedTime, targetEmail);
		}

		private void CreateAndSendEvent(object sourceItem, string originalSubject, string originalBody, DateTime originalTime, string targetEmail)
		{
			Outlook.AppointmentItem newAppointment = null;

			try
			{
				newAppointment = (Outlook.AppointmentItem)this.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);

				newAppointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
				newAppointment.Subject = $"FW: {originalSubject}";
				newAppointment.Body = $"Original Received: {originalTime}\n\n{originalBody}";

				newAppointment.Start = DateTime.Now.AddMinutes(10);
				newAppointment.Duration = 30;

				// Attachments
				if (sourceItem != null)
				{
					CopyAttachments(sourceItem, newAppointment);
				}

				Outlook.Recipient recipient = newAppointment.Recipients.Add(targetEmail);
				recipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired;

				if (recipient.Resolve())
				{
					newAppointment.Send();
				}
				else
				{
					newAppointment.Save();
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"Error creating event: {ex.Message}");
			}
			finally
			{
				if (newAppointment != null) Marshal.ReleaseComObject(newAppointment);
			}
		}

		private void CopyAttachments(object sourceItem, Outlook.AppointmentItem targetItem)
		{
			Outlook.Attachments sourceAttachments = null;

			if (sourceItem is Outlook.MailItem mail) sourceAttachments = mail.Attachments;
			else if (sourceItem is Outlook.MeetingItem meeting) sourceAttachments = meeting.Attachments;

			if (sourceAttachments != null && sourceAttachments.Count > 0)
			{
				string tempFolder = Path.GetTempPath();

				foreach (Outlook.Attachment att in sourceAttachments)
				{
					string tempFile = null;
					try
					{
						tempFile = Path.Combine(tempFolder, att.FileName);
						att.SaveAsFile(tempFile);
						targetItem.Attachments.Add(tempFile, Outlook.OlAttachmentType.olByValue, 1, att.DisplayName);
					}
					catch { }
					finally
					{
						if (tempFile != null && File.Exists(tempFile))
						{
							try { File.Delete(tempFile); } catch { }
						}
					}
				}
			}
		}

		private bool IsLocatedInDeletedItems(object item)
		{
			try
			{
				Outlook.MAPIFolder itemParent = null;

				if (item is Outlook.MailItem m) itemParent = m.Parent as Outlook.MAPIFolder;
				else if (item is Outlook.MeetingItem mt) itemParent = mt.Parent as Outlook.MAPIFolder;

				if (itemParent == null) return false;

				Outlook.MAPIFolder deletedFolder = _nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);

				if (deletedFolder != null && itemParent.EntryID == deletedFolder.EntryID)
				{
					return true;
				}
			}
			catch
			{
				return false;
			}
			return false;
		}

		private string GetSenderAddress(Outlook.AddressEntry sender)
		{
			if (sender == null) return string.Empty;
			try
			{
				if (sender.Type == "SMTP") return sender.Address;
				if (sender.Type == "EX")
				{
					Outlook.ExchangeUser exUser = sender.GetExchangeUser();
					if (exUser != null) return exUser.PrimarySmtpAddress;
				}
				return sender.Address;
			}
			catch { return string.Empty; }
		}

		public void RunManualTest()
		{
			string targetEmail = Properties.Settings.Default.TargetEmail;
			if (string.IsNullOrEmpty(targetEmail))
			{
				System.Windows.Forms.MessageBox.Show("Please configure a receiver email first.", "Test Failed");
				return;
			}
			CreateAndSendEvent(null, "[TEST] Manual Check", "Test body.", DateTime.Now, targetEmail);
			System.Windows.Forms.MessageBox.Show($"Test invite sent to: {targetEmail}", "Test Initiated");
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			_inboxItems = null; _inbox = null; _nameSpace = null;
		}

		#region VSTO generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}
		#endregion
	}
}
