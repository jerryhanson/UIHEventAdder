using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace EventAdder
{
	public partial class ForwarderRibbon
	{
		private void ForwarderRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			// Load the saved email setting when the Ribbon appears
			try
			{
				editBoxEmail.Text = Properties.Settings.Default.TargetEmail;
			}
			catch { /* Ignore first-run errors */ }
		}

		private void btnSaveEmail_Click(object sender, RibbonControlEventArgs e)
		{
			string newEmail = editBoxEmail.Text.Trim();

			if (!string.IsNullOrEmpty(newEmail))
			{
				Properties.Settings.Default.TargetEmail = newEmail;
				Properties.Settings.Default.Save();
				MessageBox.Show($"Receiver updated to: {newEmail}", "Configuration Saved");
			}
			else
			{
				MessageBox.Show("Please enter a valid email address.", "Error");
			}
		}

		private void btnTestNow_Click(object sender, RibbonControlEventArgs e)
		{
			// Trigger the manual test in ThisAddIn
			Globals.ThisAddIn.RunManualTest();
		}

        private void btnPause_Click(object sender, RibbonControlEventArgs e)
		{
			// Check if it is a ToggleButton (Recommended)
			RibbonToggleButton toggle = sender as RibbonToggleButton;

			if (toggle != null)
			{
				if (toggle.Checked)
				{
					// USER CLICKED "PAUSE"
					// Set the pause time to 30 minutes from now
					ThisAddIn.PauseUntil = DateTime.Now.AddMinutes(30);

					toggle.Label = "Resumes: " + ThisAddIn.PauseUntil.ToShortTimeString();
					MessageBox.Show($"Forwarding is PAUSED for 30 minutes.\nIt will auto-resume at {ThisAddIn.PauseUntil.ToShortTimeString()}.", "Paused");
				}
				else
				{
					// USER CLICKED "RESUME" (Unchecked the button)
					// Reset the timer
					ThisAddIn.PauseUntil = DateTime.MinValue;

					toggle.Label = "Pause 30m";
					MessageBox.Show("Forwarding Resumed.", "Active");
				}
			}
		}
	}
}
