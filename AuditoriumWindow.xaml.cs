/////////////////////////////////////////////////////////////////////////////////////////
//Author: Adam Kim
//Created On: 3/2/2022
//Last Modified On: 8/16/2022
//Copyright: USAF // JT4 LLC
//Description: Secondary window of the PaperLoad application
/////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PaperLoad_WPF
{
	public partial class AuditoriumWindow : Window
	{
		/// <summary>
		/// Main function that initializes the secondary Window
		/// </summary>
		public AuditoriumWindow()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Function to check if inputed value is a integer or not. If not do not accept value
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any text composition event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void CheckIfInt(object sender, TextCompositionEventArgs e)
		{
			try
			{
				TextBox textBox = sender as TextBox;

				Regex tempRegex = new Regex("[^0-9]+");
				e.Handled = tempRegex.IsMatch(e.Text);

				if (String.IsNullOrEmpty(e.Text) || e.Text.Length < 4 || int.Parse(e.Text) > 2359)
				{
					textBox.Background = Brushes.PaleVioletRed;
				}
				else
				{
					textBox.Background = Brushes.White;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("There was a failure in the 'CheckIfInt' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(ex.ToString());
			}
		}
	}
}
