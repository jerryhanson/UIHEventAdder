using System;
using System.Runtime.InteropServices;
using System.IO; // Required for file handling
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EventAdder
{
	public partial class ThisAddIn
	{
		private Outlook.NameSpace _nameSpace;
		private Outlook.MAPIFolder _inbox;
		private Outlook.Items _inboxItems;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			_nameSpace = this.Application.GetNamespace("MAPI");
			_inbox = _nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
			_inboxItems = _inbox.Items;

			// Monitor the Inbox
			_inboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(InboxItems_ItemAdd);
		}

		private void InboxItems_ItemAdd(object item)
		{
			try
			{
				if (item is Outlook.MailItem mail)
				{
					// Pass the 'mail' object as the source
					ProcessItem(mail, mail.Sender, mail.Subject, mail.Body, mail.SentOn);
				}
				else if (item is Outlook.MeetingItem meeting)
				{
					// Use dynamic to safely access Sender on MeetingItems
					dynamic dMeeting = meeting;
					Outlook.AddressEntry sender = dMeeting.Sender;

					// Pass the 'meeting' object as the source
					ProcessItem(meeting, sender, meeting.Subject, meeting.Body, meeting.SentOn);
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"Error processing item: {ex.Message}");
			}
		}

		// UPDATED: Now accepts 'object sourceItem'
		private void ProcessItem(object sourceItem, Outlook.AddressEntry sender, string subject, string body, DateTime receivedTime)
		{
			if (sender == null) return;

			string targetEmail = Properties.Settings.Default.TargetEmail;
			if (string.IsNullOrEmpty(targetEmail)) return;

			string senderAddress = GetSenderAddress(sender);

			// Self-Exclusion Check
			if (string.Equals(senderAddress, targetEmail, StringComparison.OrdinalIgnoreCase))
			{
				return;
			}

			CreateAndSendEvent(sourceItem, subject, body, receivedTime, targetEmail);
		}

		// UPDATED: Now accepts 'object sourceItem' and handles attachments
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

				// --- NEW: ATTACHMENT HANDLING ---
				if (sourceItem != null)
				{
					CopyAttachments(sourceItem, newAppointment);
				}
				// --------------------------------

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

		// NEW HELPER METHOD: Copies attachments via Temp folder
		private void CopyAttachments(object sourceItem, Outlook.AppointmentItem targetItem)
		{
			Outlook.Attachments sourceAttachments = null;

			// Determine if source is Mail or Meeting
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
						// 1. Generate a temp path
						// We use a simplified name to avoid path too long errors, 
						// but you could use att.FileName if preferred.
						tempFile = Path.Combine(tempFolder, att.FileName);

						// 2. Save attachment to disk
						att.SaveAsFile(tempFile);

						// 3. Attach file to the new appointment
						targetItem.Attachments.Add(tempFile, Outlook.OlAttachmentType.olByValue, 1, att.DisplayName);
					}
					catch
					{
						// Skip attachments that fail to save (e.g. embedded OLE objects)
					}
					finally
					{
						// 4. Cleanup temp file
						if (tempFile != null && File.Exists(tempFile))
						{
							try { File.Delete(tempFile); } catch { }
						}
					}
				}
			}
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

			// Pass 'null' for sourceItem since this is a manual test with no real email source
			CreateAndSendEvent(null, "[TEST] Manual Check", "Test body.", DateTime.Now, targetEmail);
			System.Windows.Forms.MessageBox.Show($"Test invite sent to: {targetEmail}", "Test Initiated");
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			_inboxItems = null;
			_inbox = null;
			_nameSpace = null;
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