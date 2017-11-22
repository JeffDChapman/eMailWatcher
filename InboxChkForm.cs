using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Microsoft.Office.Interop.Outlook;


namespace InboxCheck
{

	public class InboxChkForm : System.Windows.Forms.Form
	{
		private Microsoft.Office.Interop.Outlook._Application OutlookApp = null;
		private Microsoft.Office.Interop.Outlook.MAPIFolder inbox = null;
		private System.Windows.Forms.Timer tmrUpdate;
		private System.Windows.Forms.Timer tmrShowOutlook;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Timer tmrColorSpin;
		int rval = 0;
		int gval = 0;
		int bval = 0;
		int spinflag = 0;

		public InboxChkForm()
		{
			InitializeComponent();

			OutlookApp = new Microsoft.Office.Interop.Outlook.ApplicationClass();
			if (OutlookApp != null)
			{ 
				inbox = OutlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox); 
				this.Text = inbox.UnReadItemCount.ToString();
			}
		}


		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code


		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InboxChkForm));
            this.tmrUpdate = new System.Windows.Forms.Timer(this.components);
            this.tmrShowOutlook = new System.Windows.Forms.Timer(this.components);
            this.tmrColorSpin = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // tmrUpdate
            // 
            this.tmrUpdate.Enabled = true;
            this.tmrUpdate.Interval = 60000;
            this.tmrUpdate.Tick += new System.EventHandler(this.tmrUpdate_Tick);
            // 
            // tmrShowOutlook
            // 
            this.tmrShowOutlook.Tick += new System.EventHandler(this.tmrShowOutlook_Tick);
            // 
            // tmrColorSpin
            // 
            this.tmrColorSpin.Enabled = true;
            this.tmrColorSpin.Tick += new System.EventHandler(this.tmrColorSpin_Tick);
            // 
            // InboxChkForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(116, 7);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InboxChkForm";
            this.Text = "0";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Activated += new System.EventHandler(this.InboxChkForm_Activated);
            this.Click += new System.EventHandler(this.InboxChkForm_Click);
            this.Resize += new System.EventHandler(this.InboxChkForm_Resize);
            this.ResumeLayout(false);

		}
		#endregion

		[STAThread]
		static void Main() 
		{
			System.Windows.Forms.Application.Run(new InboxChkForm());
		}

		private void tmrUpdate_Tick(object sender, System.EventArgs e)
		{
			int MailCountHold = inbox.UnReadItemCount;
			if (MailCountHold > 0)
				{this.WindowState = FormWindowState.Normal;}
			else
				{this.WindowState = FormWindowState.Minimized;}
			this.Text = inbox.UnReadItemCount.ToString();
		}

		private void InboxChkForm_Activated(object sender, System.EventArgs e)
		{
            //if (this.WindowState == FormWindowState.Normal)
            //{
            //    this.tmrShowOutlook.Enabled = true;
            //}
		}

		private void tmrShowOutlook_Tick(object sender, System.EventArgs e)
		{
			this.tmrShowOutlook.Enabled = false;
			this.WindowState = FormWindowState.Minimized;
			System.Diagnostics.Process.Start("outlook.exe", "/recycle");
		}

		private void InboxChkForm_Resize(object sender, System.EventArgs e)
		{
			if (this.WindowState == FormWindowState.Maximized)
			{
				this.WindowState = FormWindowState.Normal;
				this.tmrShowOutlook.Enabled = true;
			}
		}

		private void InboxChkForm_Click(object sender, System.EventArgs e)
		{
			this.tmrShowOutlook.Enabled = true;
		}

		private void tmrColorSpin_Tick(object sender, System.EventArgs e)
		{
			switch(spinflag)       
			{         
				case 0:   
					if (rval < 255) {rval += 1;}
					else {if (gval < 255) {gval += 1;}
						else {if (bval < 255) {bval += 1;}
								else {gval = 0; bval = 0; rval = 0; spinflag = 1;}}
					}
					break;                  
				case 1:            
					if (gval < 255) {gval += 1;}
					else {if (bval < 255) {bval += 1;}
						  else {if (rval < 255) {rval += 1;}
								else {gval = 0; bval = 0; rval = 0; spinflag = 2;}}
					}
					break;
				case 2:            
					if (bval < 255) {bval += 1;}
					else {if (rval < 255) {rval += 1;}
						  else {if (gval < 255) {gval += 1;}
								else {gval = 0; bval = 0; rval = 0; spinflag = 0;}}
					}
					break;
				default:            
					break;      
			}


			this.BackColor = Color.FromArgb(rval, gval, bval);
			this.Refresh();
		}

	}
}
