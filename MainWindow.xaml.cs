/////////////////////////////////////////////////////////////////////////////////////////
//Author: Adam Kim
//Created On: 3/2/2022
//Last Modified On: 8/16/2022
//Copyright: USAF // JT4 LLC
//Description: Main window of the PaperLoad application
/////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace PaperLoad_WPF
{
	/// <summary>
	/// Class for generating a user input dialog box. Specifically sized for prompting for the half sheet initialed approval, but can be used for any user input.
	/// </summary>
	public static class Prompt
	{
		/// <summary>
		/// Function to present a dialog prompt where a user can input a value which is returned.
		/// </summary>
		/// <param name="label">The text found in the body of the dialog box</param>
		/// <param name="caption">The caption found at the top of the dialog box</param>
		/// <param name="text">The text found in the body of the text box</param>
		/// <param name="formWidth">The width of the form</param>
		/// <param name="formHeight">The height of the form</param>
		/// <param name="boxWidth">The width of the label and textbox</param>
		/// <param name="boxHeight">The height of the label and textbox</param>
		/// <returns>A string of the user inputted value</returns>
		public static string ShowDialog(string label, string caption, string text, int formWidth, int formHeight, int boxWidth, int boxHeight)
		{
			System.Windows.Forms.Form prompt = new System.Windows.Forms.Form()
			{
				Width = formWidth,
				Height = formHeight,
				FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
				Text = caption,
				StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
			};

			System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label() { Left = 35, Top = 15, Width = boxWidth, Height = 60, Text = label };
			System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Text = text, Left = 35, Top = 50, Width = boxWidth, Height = boxHeight, MaxLength = 435, WordWrap = true, Multiline = true };
			System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button() { Text = "OK", Left = boxWidth - 15, Top = boxHeight + 58, Width = 50, DialogResult = System.Windows.Forms.DialogResult.OK };
			confirmation.Click += (sender, e) => { prompt.Close(); };
			prompt.Controls.Add(textBox);
			prompt.Controls.Add(confirmation);
			prompt.Controls.Add(textLabel);
			prompt.AcceptButton = confirmation;

			return prompt.ShowDialog() == System.Windows.Forms.DialogResult.OK ? textBox.Text : "";
		}
	}

	public partial class MainWindow : Window
	{
		/// <summary>
		/// Global variables
		/// </summary>
		DataRelation paperloadDR;
		int missionNumber = 1;
		int playerNumber = 1;
		DataSet paperloadDS = new DataSet("PaperLoad_Dataset");
		DataTable aircraftDT = new DataTable("Aircraft_Datatable");
		DataTable sortieDT = new DataTable("Sortie_Datatable");
		string unprocessedFileDirectory = Convert.ToString(Directory.CreateDirectory(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_LMT Files\\Unprocessed"));
		string processedFileDirectory = Convert.ToString(Directory.CreateDirectory(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_LMT Files\\Processed"));
		string rampodFileDirectory = Convert.ToString(Directory.CreateDirectory(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_RAMPOD Exports"));
		string ndcarFileDirectory = Convert.ToString(Directory.CreateDirectory(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_NDCar Exports"));
		string unprocessedFileName = Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_LMT Files\\Unprocessed\\" + DateTime.Today.ToString("d-MMM-yy").ToUpper() + ".xml");
		string halfsheetTemplateFileName = Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location)) + "\\assets\\HalfSheetTemplate.xlsx";

		/// <summary>
		/// Main function that initializes the main window and calls the CreateFile function
		/// </summary>
		/// <returns>None (Void)</returns>
		public MainWindow()
		{
			try
			{
				InitializeComponent();
				CreateFile();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'MainWindow' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to aircraft to datagrid/datatable and sortie data to datagrid/datatable if it hasn't been added
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void AddAircraft_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				string tempFields = EmptyFields(false);

				// Check if all necessary fields have been filled in
				if (!String.IsNullOrEmpty(tempFields))
				{
					MessageBox.Show(Application.Current.MainWindow, "Please fill out or modify the following fields:" + tempFields, "ERROR");
				}
				// Check if duplicate IFFs and Pod Serials before submitting record
				else if (!DuplicateEntries())
				{
					// Add row to sortie datatable if not added and not editing aircraft
					if (UpdateRowCount() == 0 && aircraftAddSortie_button.IsEnabled)
					{
						sortieDT.Rows.Add(new Object[]
						{
							missionNumber,																	// Sortie Mission Number
							false,																			// Sortie Is Mission Submitted
							GenerateMissionID(),															// Sortie Mission ID
							sortieDate_input.Text,															// Sortie Date
							ConvertRangeTimes(true),														// Sortie Range Start Time
							ConvertRangeTimes(false),														// Sortie Range End Time
							sortieProject_input.Text.Replace(" ", String.Empty).ToUpper(),                  // Sortie Project Number
							int.Parse(sortieNumCD_input.Text.Replace(" ", String.Empty)),					// Sortie Number of CDs
							RecordedStations(),																// Sortie Recorded Stations
							0,																				// Sortie HAA Count
							0,																				// Sortie LAA Count
							String.Empty,																	// Sortie Notes
							String.Empty																	// Sortie Auditoriums
						});
					}

					// Always add row to aircraft datatable
					aircraftDT.Rows.Add(new Object[]
					{
						missionNumber,																					// Aircraft Mission Number
						playerNumber,																					// Aircraft Player Number
						false,																							// Aircraft Is Player Submitted
						aircraftLowAct_input.IsChecked,																	// Aircraft Is Low Activity
						aircraftUnit_input.Text,																		// Aircraft Unit
						new string(aircraftCallsign_input.Text.Where(c => char.IsLetter(c)).ToArray()).ToUpper(),       // Aircraft Callsign
						aircraftType_input.Text,																		// Aircraft Type
						aircraftStation_input.Text,																		// Aircraft Station
						aircraftTailNum_input.Text.Replace(" ", String.Empty).ToUpper(),								// Aircraft Tail Number
						int.Parse(aircraftIFF_input.Text.Replace(" ", String.Empty)),									// Aircraft IFF
						aircraftPodSerial_input.Text.Replace(" ", String.Empty),										// Aircraft Pod Serial
						aircraftTrackStatus_input.Text																	// Aircraft Track Status
					});

					// Accept changes and write to XML file
					sortieDG.UnselectAll();
					aircraftDG.UnselectAll();
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);

					// Function calls to removed selected pod serial, clear fields, re-enable buttons, and update row count
					RemovePodSerialNumber(aircraftPodSerial_input.Text);
					AircraftClearFields();
					DisableButtons(false, false);
					UpdateRowCount();
					
					// Set player number
					SetMaxPlayerNumber();

					// Set focus to first aircraft field
					aircraftLowAct_input.Focus();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AddAircraft_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to add a pod serial number back to the respective dropdown list and also sorts said list
		/// </summary>
		/// <param name="serial">String of the pod serial to add back to the dropdown list</param>
		/// <returns>None (Void)</returns>
		private void AddPodSerialNumber(string serial)
		{
			try
			{
				// Create temp list to store values in
				List<string> tempList = new List<string>();

				// Check if LAA
				if (!serial.Equals("N/A"))
				{
					// Add current pod serial back to dropdown list
					aircraftPodSerial_input.Items.Add(serial);

					// Add all available pod serials to the temp list
					foreach (string tempItem in aircraftPodSerial_input.Items)
					{
						tempList.Add(tempItem);
					}

					// Sort the temp list and clear the existing dropdown list
					tempList.Sort();
					aircraftPodSerial_input.Items.Clear();

					// Add all of the values in the temp list to the dropdown list
					foreach (string tempString in tempList)
					{
						aircraftPodSerial_input.Items.Add(tempString);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AddPodSerialNumber' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to update sortie in datagrid/datatable
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void AddSortie_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (UpdateRowCount() != 0)
				{
					foreach (DataRow sortieRow in sortieDT.Rows)
					{
						if ((int)sortieRow["Sortie_MissionNumber"] == missionNumber)
						{
							sortieRow["Sortie_IsMissionSubmitted"] = true;
							sortieRow["Sortie_HAACount"] = Convert.ToInt32(rowCountHAA_label.Content);
							sortieRow["Sortie_LAACount"] = Convert.ToInt32(rowCountLAA_label.Content);
							break;
						}
					}

					// ChangePlayerSubmitted accepts and writes changes to XML file
					ChangePlayersSubmitted(missionNumber, true);

					// Set mission number
					SetMaxMissionNumber();

					// Function calls to clear fields, change submitted status, rebuild the pod serial list, and update the row count
					AircraftClearFields();
					SortieClearFields();
					RebuildPodSerialList();
					UpdateRowCount();

					// Set focus to first sortie input field
					sortieMissionIDTime_input.Focus();
				}
				else
				{
					MessageBox.Show(Application.Current.MainWindow, "No aircraft(s) to add to sortie.", "ERROR");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AddSortie_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to clear all needed fields after submitting an aircraft
		/// </summary>
		/// <returns>None (Void)</returns>
		private void AircraftClearFields()
		{
			try
			{
				// If aircraft is low activity then change text to 'N/A'
				if ((bool)aircraftLowAct_input.IsChecked)
				{
					aircraftTailNum_input.Text = "N/A";
					aircraftPodSerial_input.Text = "N/A";
				}
				else
				{
					aircraftTailNum_input.Text = String.Empty;
					aircraftPodSerial_input.Text = String.Empty;
				}
				
				// Clear these fields regardless of condition
				aircraftIFF_input.Text = String.Empty;
				aircraftTrackStatus_input.Text = String.Empty;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AircraftClearFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to fill all fields with the values from the double clicked aircraft from the aircraft datagrid
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any mouse button event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void AircraftDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			try
			{
				if (aircraftDG.SelectedItem != null && aircraftAddSortie_button.IsEnabled)
				{
					// Get currently selected row
					DataRowView selectedRow = aircraftDG.SelectedItem as DataRowView;
					DataRow[] aircraftRow = aircraftDT.Select("Aircraft_PlayerNumber = " + selectedRow.Row[1].ToString());

					// Populate fields with data from DataRow
					aircraftLowAct_input.IsChecked = (bool)aircraftRow[0]["Aircraft_IsLowActivity"];

					// Change input mode based on populated data
					AircraftInputModeChange();

					aircraftUnit_input.Text = aircraftRow[0]["Aircraft_Unit"].ToString();
					aircraftCallsign_input.Text = aircraftRow[0]["Aircraft_Callsign"].ToString();
					aircraftType_input.Text = aircraftRow[0]["Aircraft_Type"].ToString();
					aircraftStation_input.Text = aircraftRow[0]["Aircraft_Station"].ToString();
					aircraftTailNum_input.Text = aircraftRow[0]["Aircraft_TailNumber"].ToString();
					aircraftIFF_input.Text = aircraftRow[0]["Aircraft_IFF"].ToString();
					aircraftTrackStatus_input.Text = aircraftRow[0]["Aircraft_TrackStatus"].ToString();

					// Set player number
					playerNumber = (int)aircraftRow[0]["Aircraft_PlayerNumber"];

					// Add current pod serial back to dropdown list and select it
					AddPodSerialNumber(aircraftRow[0]["Aircraft_PodSerialNumber"].ToString());
					aircraftPodSerial_input.SelectedValue = aircraftRow[0]["Aircraft_PodSerialNumber"];

					// Delete the correct row from the aircraft datatable, accept changes, and save XML file
					aircraftRow[0].Delete();
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);

					// Function calls to disable buttons and update row count
					DisableButtons(true, false);
					UpdateRowCount();

					// Hide sortieModify_input checkbox
					sortieModify_input.Visibility = Visibility.Hidden;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AircraftDataGrid_MouseDoubleClick' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to delete a selected aircraft whenever the "Delete Row" option is selected in the aircraft datagrid context menu
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void AircraftDataGridContextDelete_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (MessageBox.Show(Application.Current.MainWindow, "Are you sure you want to delete this aircraft?", "WARNING", MessageBoxButton.YesNo) == MessageBoxResult.Yes && aircraftDG.SelectedItem != null)
				{
					// Get currently selected row
					DataRowView selectedRow = aircraftDG.SelectedItem as DataRowView;
					DataRow[] aircraftRow = aircraftDT.Select("Aircraft_PlayerNumber = " + selectedRow.Row[1].ToString());

					// Add current pod serial back to dropdown list
					AddPodSerialNumber(aircraftRow[0]["Aircraft_PodSerialNumber"].ToString());

					// Delete the correct row from the aircraft datatable, accept changes, and save XML file
					aircraftRow[0].Delete();
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);

					// Update the row count in the status bar
					UpdateRowCount();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AircraftDataGridContextDelete_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function for when the input mode is changed (HAA/LAA). Broken out from event function so it can be called without params
		/// </summary>
		/// <returns>None (Void)</returns>
		private void AircraftInputModeChange()
		{
			try
			{
				// Setup Aircraft Types list to use
				List<string> aircraftTypes = new List<string>("A-4,A-10,AV-8,B-52H,F-15,F-16,F-18,F-18EF,L-159,M2000D,TORNADO".Split(',').ToList());

				// Setup Aircraft Track Statuses list to use
				List<string> aircraftTrackStatuses = new List<string>("GT,BT,NT,CNX".Split(',').ToList());

				if ((bool)aircraftLowAct_input.IsChecked)
				{
					// Fill input boxes with dummy data
					aircraftTailNum_input.Text = "N/A";
					aircraftStation_input.Text = "N/A";
					aircraftPodSerial_input.Text = "N/A";

					// Add LAA
					aircraftTypes.AddRange("B-1B,B-2,C-130,C-17A,CV-22A,E-2C,E-3,F-22A,F-35A,F-35B,F-35C,F-5,H-60,KC-135,MF-1,MQ-9,T-38".Split(',').ToList());

					// Remove HAA statuses
					aircraftTrackStatuses.RemoveRange(1, 2);
				}
				else
				{
					// Empty input boxes
					aircraftTailNum_input.Text = String.Empty;
					aircraftStation_input.Text = String.Empty;
					aircraftPodSerial_input.Text = String.Empty;
				}

				// Clear all Aircraft Types then add values from aircraftTypes list
				aircraftType_input.Items.Clear();
				foreach (string aircraftType in aircraftTypes)
				{
					aircraftType_input.Items.Add(aircraftType);
				}

				// Clear all Aircraft Track Statuses then add values from aircraftTrackStatuses list
				aircraftTrackStatus_input.Items.Clear();
				foreach (string aircraftTrackStatus in aircraftTrackStatuses)
				{
					aircraftTrackStatus_input.Items.Add(aircraftTrackStatus);
				}

				// Hack to fix statueses bug which causes the dropdown to be incorrectly sized
				aircraftTrackStatus_input.IsDropDownOpen = true;
				aircraftTrackStatus_input.IsDropDownOpen = false;

				// Enable/Disable inputs based on current mode
				aircraftStation_input.IsEnabled = !(bool)aircraftLowAct_input.IsChecked;
				aircraftTailNum_input.IsEnabled = !(bool)aircraftLowAct_input.IsChecked;
				aircraftPodSerial_input.IsEnabled = !(bool)aircraftLowAct_input.IsChecked;

				// Since the units are blanked in nested function call we need to manually reassign the value to the field
				string tempCallsign = aircraftCallsign_input.Text;
				SelectAircraftFromUnit();
				aircraftCallsign_input.Text = tempCallsign;
				SelectAircraftFromCallsign();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AircraftInputModeChange' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Event function to call the broken out function when the input mode is changed (HAA/LAA)
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void AircraftInputModeChange_Click(object sender, EventArgs e)
		{
			try
			{
				AircraftInputModeChange();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'AircraftInputModeChange_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the available callsign options in the drop down menu. Changes values based on the selected input in the aircraftUnit_input text field
		/// </summary>
		/// <returns>None (Void)</returns>
		private void ChangeCallsignOptions()
		{
			try
			{
				// Create list to use for Aircraft Callsigns
				List<string> aircraftCallsigns = new List<string>();

				// Add units dependant on unit
				if (aircraftUnit_input.Text == "16 WPS")
				{
					aircraftCallsigns.AddRange("SNAKE,COBRA,WOLF,PYTHON,WEASEL".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "17 WPS")
				{
					aircraftCallsigns.AddRange("HOSS,HOOTR".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "6 WPS")
				{
					aircraftCallsigns.AddRange("BONG,SCAT,GRAVE,PAPPY,SHOCK,SKULL".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "64 AGRS")
				{
					aircraftCallsigns.AddRange("MIG,IVAN,GOMER,DRAGO".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "65 AGRS")
				{
					aircraftCallsigns.AddRange("DRAGON".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "66 WPS")
				{
					aircraftCallsigns.AddRange("HOG,CANNON,GUNN,RIFLE,SANDY,NAIL".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "422 TES")
				{
					aircraftCallsigns.AddRange("VIPER,VENOM,STRIKE,RAPTOR,BOLT,BOAR,EAGLE".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "433 WPS")
				{
					aircraftCallsigns.AddRange("SATAN,DEMON,RAMBO,CONAN".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "34 WPS")
				{
					aircraftCallsigns.AddRange("STING,ROYAL".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "26 WPS")
				{
					aircraftCallsigns.AddRange("SAVAGE,MUSTANG".Split(',').ToList());
				}
				else if (aircraftUnit_input.Text == "TOP ACES")
				{
					aircraftCallsigns.AddRange("ACES".Split(',').ToList());
				}

				// Clear existing units and add ones from list
				aircraftCallsign_input.Items.Clear();
				foreach (string aircraftCallsign in aircraftCallsigns)
				{
					aircraftCallsign_input.Items.Add(aircraftCallsign);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ChangeCallsignOptions' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the Aircraft_IsPlayerSubmitted value for the specified mission number in the XML file
		/// </summary>
		/// <param name="missionNumber">Integer of the specified mission number</param>
		/// <param name="changedValue">Boolean of the value to change General_PlayersSubmitted to</param>
		/// <returns>None (Void)</returns>
		private void ChangePlayersSubmitted(int missionNumber, bool changedValue)
		{
			try
			{
				// Find each row in the datatable with the specified missionNumber and change Aircraft_IsPlayerSubmitted to specified value
				foreach (DataRow aircraftRow in aircraftDT.Rows)
				{
					if ((int)aircraftRow["Aircraft_MissionNumber"] == missionNumber)
					{
						aircraftRow["Aircraft_IsPlayerSubmitted"] = changedValue;
					}
				}

				// Accept and write changes to XML file
				paperloadDS.AcceptChanges();
				paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ChangePlayersSubmitted' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the available station options in the drop down menu. Changes values based on the selected input in the aircraftType_input text field
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void ChangeStationOptions(object sender, EventArgs e)
		{
			try
			{
				// Create list to use for Aircraft Stations
				List<string> aircraftStations = new List<string>();

				// Add stations dependant on aircraft type
				if (aircraftType_input.Text == "A-10")
				{
					aircraftStations.AddRange("1I,1O,11I,11O".Split(',').ToList());
				}
				else if (aircraftType_input.Text == "F-15")
				{
					aircraftStations.AddRange("2I,2O,8I,8O".Split(',').ToList());
				}
				else if (aircraftType_input.Text == "F-16")
				{
					aircraftStations.AddRange("1,2A,8A,9".Split(',').ToList());
				}
				else if (aircraftType_input.Text == "F-18EF")
				{
					aircraftStations.AddRange("1,2,10,11".Split(',').ToList());
				}
				else if (aircraftType_input.Text == "L-159" || aircraftType_input.Text == "TORNADO" || aircraftType_input.Text == "B-52H")
				{
					aircraftStations.AddRange("R,L".Split(',').ToList());
				}
				else
				{
					aircraftStations.AddRange("1,2,3,4,5,6,7,8,9,10,11".Split(',').ToList());
				}

				// Clear existing stations and add ones from list
				aircraftStation_input.Items.Clear();
				foreach (string aircraftStation in aircraftStations)
				{
					aircraftStation_input.Items.Add(aircraftStation);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ChangeStationOptions' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
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
				Regex tempRegex = new Regex("[^0-9]+");
				e.Handled = tempRegex.IsMatch(e.Text);
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'CheckIfInt' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to auto open the combobox dropdown menu when the element receives focus
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any keyboard focus changed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void ComboBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
		{
			try
			{
				ComboBox comboBox = sender as ComboBox;
				comboBox.IsDropDownOpen = true;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ComboBox_GotKeyboardFocus' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to limit the input values in a combobox to 10 chars (unless aircraftPodSerial_input which is 5 chars)
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void ComboBox_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				ComboBox sentOBJ = sender as ComboBox;

				if (sentOBJ != null)
				{
					// Isolate the TextBox portion of the ComboBox to set MaxLength property
					TextBox sentTextBox = sentOBJ.Template.FindName("PART_EditableTextBox", sentOBJ) as TextBox;

					if (sentTextBox != null)
					{
						if (sentOBJ.Name == "aircraftPodSerial_input")
						{
							sentTextBox.MaxLength = 5;
						}
						else
						{
							sentTextBox.MaxLength = 10;
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ComboBox_Loaded' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to blank a combobox if the inputed value does not match an item in the drop down list
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void ComboBox_LostFocus(object sender, RoutedEventArgs e)
		{
			try
			{
				ComboBox comboBox = sender as ComboBox;

				// If the selected item is not avaialbe in list, the ComboBox has lost keyboard focus, and the text isn't empty then replace with empty string
				if (comboBox.SelectedIndex == -1 && !comboBox.IsKeyboardFocusWithin && !comboBox.Text.Equals(String.Empty))
				{
					MessageBox.Show(Application.Current.MainWindow, "Please select a valid option from the dropdown list.", "ERROR");
					comboBox.Text = String.Empty;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ComboBox_LostFocus' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to convert the entered range times into a DateTime formatted value (SHOULD UPDATE TO PASS BY REFERENCE TO AVOID DOUBLE FUNCTION CALL)
		/// </summary>
		/// <param name="start">Boolean value that denotes if it is calculating the range start time</param>
		/// <returns>Converted DateTime</returns>
		private DateTime? ConvertRangeTimes(bool start)
		{
			try
			{
				// Return requested value
				if (start)
				{
					return DateTime.Parse(sortieDate_input.Text).Add(TimeSpan.Parse(sortieStartTime_input.Text.Substring(0, 2) + ":" + sortieStartTime_input.Text.Substring(2, 2)));
				}
				else
				{
					// Check if end time is on next day
					if (int.Parse(sortieStartTime_input.Text) > int.Parse(sortieEndTime_input.Text))
					{
						return DateTime.Parse(sortieDate_input.Text).Add(TimeSpan.Parse(sortieEndTime_input.Text.Substring(0, 2) + ":" + sortieEndTime_input.Text.Substring(2, 2))).AddDays(1);
					}
					else
					{
						return DateTime.Parse(sortieDate_input.Text).Add(TimeSpan.Parse(sortieEndTime_input.Text.Substring(0, 2) + ":" + sortieEndTime_input.Text.Substring(2, 2)));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ConvertRangeTimes' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return null;
			}
		}

		/// <summary>
		/// Function to create the aircraft and sortie datatables, then populate with read values in XML file exists
		/// </summary>
		/// <returns>None (Void)</returns>
		private void CreateFile()
		{
			try
			{
				// Create the aircraft datatable and add to dataset
				aircraftDT.Columns.Add("Aircraft_MissionNumber", typeof(int));
				aircraftDT.Columns.Add("Aircraft_PlayerNumber", typeof(int));
				aircraftDT.Columns.Add("Aircraft_IsPlayerSubmitted", typeof(bool));
				aircraftDT.Columns.Add("Aircraft_IsLowActivity", typeof(bool));
				aircraftDT.Columns.Add("Aircraft_Unit", typeof(string));
				aircraftDT.Columns.Add("Aircraft_Callsign", typeof(string));
				aircraftDT.Columns.Add("Aircraft_Type", typeof(string));
				aircraftDT.Columns.Add("Aircraft_Station", typeof(string));
				aircraftDT.Columns.Add("Aircraft_TailNumber", typeof(string));
				aircraftDT.Columns.Add("Aircraft_IFF", typeof(int));
				aircraftDT.Columns.Add("Aircraft_PodSerialNumber", typeof(string));
				aircraftDT.Columns.Add("Aircraft_TrackStatus", typeof(string));
				aircraftDG.ItemsSource = aircraftDT.DefaultView;
				paperloadDS.Tables.Add(aircraftDT);

				// Create the sortie datatable
				sortieDT.Columns.Add("Sortie_MissionNumber", typeof(int));
				sortieDT.Columns.Add("Sortie_IsMissionSubmitted", typeof(bool));
				sortieDT.Columns.Add("Sortie_MissionID", typeof(string));
				sortieDT.Columns.Add("Sortie_Date", typeof(DateTime));
				sortieDT.Columns.Add("Sortie_StartRangeTime", typeof(DateTime));
				sortieDT.Columns.Add("Sortie_EndRangeTime", typeof(DateTime));
				sortieDT.Columns.Add("Sortie_ProjectNumber", typeof(string));
				sortieDT.Columns.Add("Sortie_NumOfCDs", typeof(int));
				sortieDT.Columns.Add("Sortie_RecordingStations", typeof(string));
				sortieDT.Columns.Add("Sortie_HAACount", typeof(int));
				sortieDT.Columns.Add("Sortie_LAACount", typeof(int));
				sortieDT.Columns.Add("Sortie_Notes", typeof(string));
				sortieDT.Columns.Add("Sortie_Auditoriums", typeof(string));
				sortieDG.ItemsSource = sortieDT.DefaultView;
				paperloadDS.Tables.Add(sortieDT);

				// Datagrid sorting
				aircraftDG.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("Aircraft_PlayerNumber", System.ComponentModel.ListSortDirection.Ascending));
				sortieDG.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("Sortie_MissionNumber", System.ComponentModel.ListSortDirection.Ascending));

				// Setup datarelation
				paperloadDR = paperloadDS.Relations.Add("MissionNumber", sortieDT.Columns["Sortie_MissionNumber"], aircraftDT.Columns["Aircraft_MissionNumber"], false);
				paperloadDR.Nested = true;

				// Run input mode change and rebuild serial list functions to populate AC type, track statuses, and pod serial fields with Strings instead of ComboBoxItems. Not it's original function
				// but it already accomplishes this job with minimal additional overhead
				AircraftInputModeChange();
				RebuildPodSerialList();

				// If file already exists then open and read file.
				if (File.Exists(unprocessedFileName))
				{
					// Read XML file
					paperloadDS.ReadXml(unprocessedFileName, XmlReadMode.ReadSchema);

					// Set max mission number and player number
					SetMaxMissionNumber();
					SetMaxPlayerNumber();

					// Check to see if sortie fields need to be repopulated
					if (UpdateRowCount() != 0)
					{
						RepopulateSortieFields();
					}
				}
				else
				{
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'CreateFile' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary> 
		/// Function to clear the current selection in the datagrid when the datagrid loses focus
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any mouse button event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void DataGrid_LostFocus(object sender, RoutedEventArgs e)
		{
			try
			{
				// Get selected row from datagrid
				DataGrid datagrid = sender as DataGrid;
				DataGridRow datagridRow = datagrid.ItemContainerGenerator.ContainerFromItem(datagrid.SelectedItem) as DataGridRow;

				// Check if the selected row is null
				if (datagridRow != null)
				{
						datagrid.SelectedItem = null;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'DataGrid_LostFocus' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to enable/disable buttons dependant on passed parameters
		/// </summary>
		/// <param name="disable">Boolean to denote whether to enable or disable all buttons except aircraftAdd_button</param>
		/// <param name="disableACButton">Boolean to denote whether to enable or disable aircraftAdd_button</param>
		/// <returns>None (Void)</returns>
		private void DisableButtons(bool disable, bool disableACButton)
		{
			try
			{
				// Enable/Disable buttons respective to passed parameter
				aircraftAdd_button.IsEnabled = !disableACButton;
				aircraftAddSortie_button.IsEnabled = !disable;
				sortieRAMPODExport_button.IsEnabled = !disable;
				openFile_button.IsEnabled = !disable;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'DisableButtons' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to check whether there are duplicate values being submitted
		/// </summary>
		/// <returns>Boolean to denote whether a duplicate value was submitted</returns>
		private bool DuplicateEntries()
		{
			try
			{
				// Create temp lists to reference for duplicates
				List<int> tempIFF = new List<int>();
				List<string> tempSerial = new List<string>();

				// Loop through sortie datatable until current mission is found
				foreach (DataRow sortieRow in sortieDT.Rows)
				{
					if ((int)sortieRow["Sortie_MissionNumber"] == missionNumber)
					{
						foreach (DataRow aircraftRow in sortieRow.GetChildRows(paperloadDR))
						{
							// Always add IFF
							tempIFF.Add((int)aircraftRow["Aircraft_IFF"]);

							// If list contains IFF then duplicate is found
							if (tempIFF.Contains(int.Parse(aircraftIFF_input.Text)))
							{
								MessageBox.Show(Application.Current.MainWindow, "Duplicate IFF's cannot be submitted.", "ERROR");
								return true;
							}

							// If not LAA
							if (!(bool)aircraftRow["Aircraft_IsLowActivity"])
							{
								tempSerial.Add(aircraftRow["Aircraft_PodSerialNumber"].ToString());

								// If list contains serial then duplicate is found
								if (tempSerial.Contains(aircraftPodSerial_input.Text))
								{
									MessageBox.Show(Application.Current.MainWindow, "Duplicate Pod Serial Numbers cannot be submitted.", "ERROR");
									return true;
								}
							}
						}

						break;
					}
				}

				// If no result was found in loops then return false
				return false;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'DuplicateEntries' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return true;
			}
		}

		/// <summary>
		/// Function that returns the respective empty fields
		/// /// </summary>
		/// <param name="sortieOnly">Boolean to denote if the sortie fields are the only fields to be checked</param>
		/// <returns>String of the empty fields when attempting to submit a record</returns>
		private string EmptyFields(bool sortieOnly)
		{
			try
			{
				// Create temp string to return.
				string tempString = String.Empty;

				// Check for empty fields
				if (String.IsNullOrEmpty(sortieMissionIDJulian_input.Text) || String.IsNullOrEmpty(sortieMissionIDTime_input.Text) || sortieMissionIDTime_input.Text.Length < 4 
					|| String.IsNullOrEmpty(sortieMissionIDSub_input.Text))
				{
					tempString = tempString + " [Sortie Mission ID]";
				}
				if (String.IsNullOrEmpty(sortieDate_input.Text))
				{
					tempString = tempString + " [Sortie Date]";
				}
				if (String.IsNullOrEmpty(sortieStartTime_input.Text) || sortieStartTime_input.Text.Length < 4 || int.Parse(sortieStartTime_input.Text) > 2359)
				{
					tempString = tempString + " [Range Start Time]";
				}
				if (String.IsNullOrEmpty(sortieEndTime_input.Text) || sortieEndTime_input.Text.Length < 4 || int.Parse(sortieEndTime_input.Text) > 2359)
				{
					tempString = tempString + " [Range End Time]";
				}
				if (String.IsNullOrEmpty(sortieProject_input.Text))
				{
					tempString = tempString + " [Project Number]";
				}
				if (String.IsNullOrEmpty(sortieNumCD_input.Text))
				{
					tempString = tempString + " [Number of CDs]";
				}
				if (String.IsNullOrEmpty(RecordedStations()))
				{
					tempString = tempString + " [Recorded Stations]";
				}

				// If only checking for sortie information then skip
				if (!sortieOnly)
				{
					if (String.IsNullOrEmpty(aircraftUnit_input.Text))
					{
						tempString = tempString + " [Aircraft Unit]";
					}
					if (String.IsNullOrEmpty(aircraftCallsign_input.Text))
					{
						tempString = tempString + " [Aircraft Callsign]";
					}
					if (String.IsNullOrEmpty(aircraftType_input.Text))
					{
						tempString = tempString + " [Aircraft Type]";
					}
					if (String.IsNullOrEmpty(aircraftStation_input.Text))
					{
						tempString = tempString + " [Aircraft Station]";
					}
					if (String.IsNullOrEmpty(aircraftTailNum_input.Text))
					{
						tempString = tempString + " [Aircraft Tail Number]";
					}
					if (String.IsNullOrEmpty(aircraftIFF_input.Text))
					{
						tempString = tempString + " [Aircraft IFF]";
					}
					if (String.IsNullOrEmpty(aircraftPodSerial_input.Text))
					{
						tempString = tempString + " [Aircraft Pod Serial]";
					}
					if (String.IsNullOrEmpty(aircraftTrackStatus_input.Text))
					{
						tempString = tempString + " [Aircraft Track Status]";
					}
				}

				return tempString;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'EmptyFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return String.Empty;
			}
		}

		/// <summary>
		/// Function to export sortie into half sheet
		/// </summary>
		/// <returns>None (Void)</returns>
		private void ExportSortie_HalfSheet()
		{
			try
			{
				// Get currently selected row
				DataRowView selectedSortie = sortieDG.SelectedItem as DataRowView;
				DataRow[] selectedRow = sortieDT.Select("Sortie_MissionNumber = " + selectedSortie.Row[0].ToString());

				// Check if file is already open
				if (!IsFileOpen(ndcarFileDirectory + "\\" + selectedRow[0]["Sortie_MissionID"] + ".xlsx"))
				{
					// Prompt and collect initials for half sheet
					string userInput = Prompt.ShowDialog("Please initial Half Sheet for approval (Seperate with commas).", "APPROVE HALF SHEET", String.Empty, 300, 155, 215, 22);

					if (!String.IsNullOrEmpty(userInput))
					{
						// Set up temp variables to use
						int lowCount = 0;
						int lowCNX = 0;
						int highCount = 0;
						int highCNX = 0;
						int highNonEffective = 0;
						bool flag16WPS = false;
						bool flag17WPS = false;
						bool flag6WPS = false;
						bool flag64AGRS = false;
						bool flag65AGRS = false;
						bool flag66WPS = false;
						bool flag422TES = false;
						bool flag433WPS = false;
						bool flagTOPACES = false;
						bool flag26WPS = false;
						bool flag34WPS = false;
						bool flagTDY1 = false;
						bool flagTDY2 = false;
						bool flagTDY3 = false;
						bool flagTDY4 = false;
						bool flagTDY5 = false;
						string callsigns16WPS = "16 WPS: ";
						string callsigns17WPS = "17 WPS: ";
						string callsigns6WPS = "6 WPS: ";
						string callsigns64AGRS = "64 AGRS: ";
						string callsigns65AGRS = "65 AGRS: ";
						string callsigns66WPS = "66 WPS: ";
						string callsigns422TES = "422 TES: ";
						string callsigns433WPS = "433 WPS: ";
						string callsigns34WPS = "34 WPS: ";
						string callsigns26WPS = "26 WPS: ";
						string callsignsTOPACES = "TOP ACES: ";
						string callsignsTDY1 = "TDY: ";
						string callsignsTDY2 = "TDY: ";
						string callsignsTDY3 = "TDY: ";
						string callsignsTDY4 = "TDY: ";
						string callsignsTDY5 = "TDY: ";
						string unitCallsignsLine1 = String.Empty;
						string unitCallsignsLine2 = String.Empty;
						string unitCallsignsLine3 = String.Empty;
						string unitCallsignsLine4 = String.Empty;
						string unitCallsignsLine5 = String.Empty;
						List<string> unitsList = new List<string>();
						List<string> callsignsList = new List<string>();
						List<string> callsignsListTDY = new List<string>();

						// Variables for Excel application, workbook, and worksheet
						Excel.Application excelApp = new Excel.Application();
						Excel.Workbook excelHalfSheetWorkBook = excelApp.Workbooks.Open(halfsheetTemplateFileName);
						Excel.Worksheet excelHalfSheetWorkSheet = excelHalfSheetWorkBook.Worksheets["Half Sheet"];

						// Check all records in sortie datatable for selected mission number
						foreach (DataRow sortieRow in sortieDT.Rows)
						{
							if ((int)sortieRow["Sortie_MissionNumber"] == (int)selectedRow[0]["Sortie_MissionNumber"])
							{
								// Enter half sheet header information
								excelHalfSheetWorkSheet.Cells[2, 2] = sortieRow["Sortie_MissionID"].ToString();
								excelHalfSheetWorkSheet.Cells[2, 10] = ((DateTime)sortieRow["Sortie_Date"]).ToString("d-MMM-yy");
								excelHalfSheetWorkSheet.Cells[4, 3] = ((DateTime)sortieRow["Sortie_StartRangeTime"]).ToString("HHmm") + "-" + ((DateTime)sortieRow["Sortie_EndRangeTime"]).ToString("HHmm");
								excelHalfSheetWorkSheet.Cells[4, 5] = sortieRow["Sortie_ProjectNumber"].ToString();
								excelHalfSheetWorkSheet.Cells[4, 7] = sortieRow["Sortie_NumOfCDs"].ToString();
								excelHalfSheetWorkSheet.Cells[4, 9] = sortieRow["Sortie_RecordingStations"].ToString();

								// Loop through all child rows of selected mission number
								foreach (DataRow aircraftRow in sortieRow.GetChildRows(paperloadDR))
								{
									// Check if low activity
									if ((bool)aircraftRow["Aircraft_IsLowActivity"])
									{
										// Set custom values based on Aircraft_TrackStatus
										if (aircraftRow["Aircraft_TrackStatus"].Equals("CNX"))
										{
											lowCNX = lowCNX + 1;
										}

										// Increment total LAA count
										lowCount = lowCount + 1;
									}
									else
									{
										// Set custom values based on Aircraft_TrackStatus
										if (aircraftRow["Aircraft_TrackStatus"].Equals("CNX"))
										{
											highCNX = highCNX + 1;
										}
										else if (aircraftRow["Aircraft_TrackStatus"].Equals("BT") || aircraftRow["Aircraft_TrackStatus"].Equals("NT"))
										{
											// Check if number of non-effective aircraft exceeds the non-effective table size
											if (highNonEffective < 10)
											{
												// Add non-effective information to non-effective table on half sheet
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 2] = aircraftRow["Aircraft_Unit"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 3] = aircraftRow["Aircraft_Callsign"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 4] = aircraftRow["Aircraft_Type"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 5] = aircraftRow["Aircraft_Station"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 6] = aircraftRow["Aircraft_TailNumber"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 7] = aircraftRow["Aircraft_PodSerialNumber"].ToString();
												excelHalfSheetWorkSheet.Cells[20 + highNonEffective, 8] = aircraftRow["Aircraft_TrackStatus"].ToString();
											}
											else
											{
												MessageBox.Show(Application.Current.MainWindow, "Cannot add all Non-Effective aircraft to the Half Sheet table due to a row limitation, please manually add the non-included aircraft in the notes section.", "ERROR");
											}

											// Always increment non-effective count even when not listed
											highNonEffective = highNonEffective + 1;
										}

										// Increment total HAA count
										highCount = highCount + 1;
									}

									// Generate initial Unit/Callsigns strings
									if (aircraftRow["Aircraft_Unit"].ToString() == "16 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns16WPS = callsigns16WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag16WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "17 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns17WPS = callsigns17WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag17WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "6 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns6WPS = callsigns6WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag6WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "64 AGRS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns64AGRS = callsigns64AGRS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag64AGRS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "65 AGRS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns65AGRS = callsigns65AGRS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag65AGRS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "66 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns66WPS = callsigns66WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag66WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "422 TES" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns422TES = callsigns422TES + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag422TES = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "433 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns433WPS = callsigns433WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag433WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "34 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns34WPS = callsigns34WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag34WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "26 WPS" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsigns26WPS = callsigns26WPS + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flag26WPS = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "TOP ACES" && !callsignsList.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										callsignsTOPACES = callsignsTOPACES + aircraftRow["Aircraft_Callsign"].ToString() + "/";
										callsignsList.Add(aircraftRow["Aircraft_Callsign"].ToString());
										flagTOPACES = true;
									}
									else if (aircraftRow["Aircraft_Unit"].ToString() == "TDY" && !callsignsListTDY.Contains(aircraftRow["Aircraft_Callsign"].ToString()))
									{
										if ((callsignsTDY1 + aircraftRow["Aircraft_Callsign"].ToString() + "/").Length < 83)
										{
											callsignsTDY1 = callsignsTDY1 + aircraftRow["Aircraft_Callsign"].ToString() + "/";
											flagTDY1 = true;
										}
										else if ((callsignsTDY2 + aircraftRow["Aircraft_Callsign"].ToString() + "/").Length < 83)
										{
											callsignsTDY2 = callsignsTDY2 + aircraftRow["Aircraft_Callsign"].ToString() + "/";
											flagTDY2 = true;
										}
										else if ((callsignsTDY3 + aircraftRow["Aircraft_Callsign"].ToString() + "/").Length < 83)
										{
											callsignsTDY3 = callsignsTDY3 + aircraftRow["Aircraft_Callsign"].ToString() + "/";
											flagTDY3 = true;
										}
										else if ((callsignsTDY4 + aircraftRow["Aircraft_Callsign"].ToString() + "/").Length < 83)
										{
											callsignsTDY4 = callsignsTDY4 + aircraftRow["Aircraft_Callsign"].ToString() + "/";
											flagTDY4 = true;
										}
										else if ((callsignsTDY5 + aircraftRow["Aircraft_Callsign"].ToString() + "/").Length < 83)
										{
											callsignsTDY5 = callsignsTDY5 + aircraftRow["Aircraft_Callsign"].ToString() + "/";
											flagTDY5 = true;
										}

										callsignsListTDY.Add(aircraftRow["Aircraft_Callsign"].ToString());
									}
								}

								// Break out of sortie loop since only a single sortie's half sheet should be generated
								break;
							}
						}

						// Generate final Unit/Callsign strings and add to list
						if (flag64AGRS)
						{
							callsigns64AGRS = callsigns64AGRS.Remove(callsigns64AGRS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns64AGRS);
						}
						if (flag65AGRS)
						{
							callsigns65AGRS = callsigns65AGRS.Remove(callsigns65AGRS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns65AGRS);
						}
						if (flagTOPACES)
						{
							callsignsTOPACES = callsignsTOPACES.Remove(callsignsTOPACES.Length - 1, 1) + "; ";
							unitsList.Add(callsignsTOPACES);
						}
						if (flag16WPS)
						{
							callsigns16WPS = callsigns16WPS.Remove(callsigns16WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns16WPS);
						}
						if (flag17WPS)
						{
							callsigns17WPS = callsigns17WPS.Remove(callsigns17WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns17WPS);
						}
						if (flag66WPS)
						{
							callsigns66WPS = callsigns66WPS.Remove(callsigns66WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns66WPS);
						}
						if (flag6WPS)
						{
							callsigns6WPS = callsigns6WPS.Remove(callsigns6WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns6WPS);
						}
						if (flag433WPS)
						{
							callsigns433WPS = callsigns433WPS.Remove(callsigns433WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns433WPS);
						}
						if (flag422TES)
						{
							callsigns422TES = callsigns422TES.Remove(callsigns422TES.Length - 1, 1) + "; ";
							unitsList.Add(callsigns422TES);
						}		
						if (flag34WPS)
						{
							callsigns34WPS = callsigns34WPS.Remove(callsigns34WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns34WPS);
						}
						if (flag26WPS)
						{
							callsigns26WPS = callsigns26WPS.Remove(callsigns26WPS.Length - 1, 1) + "; ";
							unitsList.Add(callsigns26WPS);
						}
						if (flagTDY1)
						{
							callsignsTDY1 = callsignsTDY1.Remove(callsignsTDY1.Length - 1, 1) + "; ";
							unitsList.Add(callsignsTDY1);

							if (flagTDY2)
							{
								callsignsTDY2 = callsignsTDY2.Remove(callsignsTDY2.Length - 1, 1) + "; ";
								unitsList.Add(callsignsTDY2);
							}
							if (flagTDY3)
							{
								callsignsTDY3 = callsignsTDY3.Remove(callsignsTDY3.Length - 1, 1) + "; ";
								unitsList.Add(callsignsTDY3);
							}
							if (flagTDY4)
							{
								callsignsTDY4 = callsignsTDY4.Remove(callsignsTDY4.Length - 1, 1) + "; ";
								unitsList.Add(callsignsTDY4);
							}
							if (flagTDY5)
							{
								callsignsTDY5 = callsignsTDY5.Remove(callsignsTDY5.Length - 1, 1) + "; ";
								unitsList.Add(callsignsTDY5);
							}
						}

						// Iterate through unitsList and append to row that has enough space to fit within the boundaries
						foreach (string unit in unitsList)
						{
							if ((unit + unitCallsignsLine1).Length < 100)
							{
								unitCallsignsLine1 = unitCallsignsLine1 + unit;
							}
							else if ((unit + unitCallsignsLine2).Length < 100)
							{
								unitCallsignsLine2 = unitCallsignsLine2 + unit;
							}
							else if ((unit + unitCallsignsLine3).Length < 100)
							{
								unitCallsignsLine3 = unitCallsignsLine3 + unit;
							}
							else if ((unit + unitCallsignsLine4).Length < 100)
							{
								unitCallsignsLine4 = unitCallsignsLine4 + unit;
							}
							else if ((unit + unitCallsignsLine5).Length < 100)
							{
								unitCallsignsLine5 = unitCallsignsLine5 + unit;
							}
							else
							{
								MessageBox.Show(Application.Current.MainWindow, "Cannot add all Unit/Callsigns to the Half Sheet table due to a row limitation, please manually add the non-included aircraft in the notes section.", "ERROR");
							}
						}

						// Assign Unit/Callsign strings to half sheet
						excelHalfSheetWorkSheet.Cells[6, 4] = unitCallsignsLine1;
						excelHalfSheetWorkSheet.Cells[7, 4] = unitCallsignsLine2;
						excelHalfSheetWorkSheet.Cells[8, 4] = unitCallsignsLine3;
						excelHalfSheetWorkSheet.Cells[9, 4] = unitCallsignsLine4;
						excelHalfSheetWorkSheet.Cells[10, 4] = unitCallsignsLine5;

						// Assign aircraft totals to half sheet table
						excelHalfSheetWorkSheet.Cells[13, 6] = highCount.ToString();
						excelHalfSheetWorkSheet.Cells[13, 7] = lowCount.ToString();
						excelHalfSheetWorkSheet.Cells[14, 6] = highCNX.ToString();
						excelHalfSheetWorkSheet.Cells[14, 7] = lowCNX.ToString();
						excelHalfSheetWorkSheet.Cells[15, 6] = highNonEffective.ToString();
						excelHalfSheetWorkSheet.Cells[16, 6] = (highCount - highCNX - highNonEffective).ToString();
						excelHalfSheetWorkSheet.Cells[16, 7] = (lowCount - lowCNX).ToString();

						// Assign notes and initials to half sheet
						excelHalfSheetWorkSheet.Cells[32, 2] = selectedRow[0]["Sortie_Notes"].ToString();
						excelHalfSheetWorkSheet.Cells[48, 9] = userInput.Replace(" ", "").ToUpper();
						
						try
						{
							// Save workbook
							excelHalfSheetWorkBook.SaveAs(ndcarFileDirectory + "\\" + selectedRow[0]["Sortie_MissionID"] + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value,
							Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive);

							// Print Worksheet. Have to include hack to not display running Excel Application triggered by Dialog.Show()
							double tempHeight = excelApp.Height;
							double tempWidth = excelApp.Width;
							excelApp.Height = 0;
							excelApp.Width = 0;

							if (excelApp.Dialogs[Excel.XlBuiltInDialog.xlDialogPrint].Show())
							{
								excelApp.Visible = false;
								excelApp.Height = tempHeight;
								excelApp.Width = tempWidth;

								// Prompt success if successful save and print
								MessageBox.Show(Application.Current.MainWindow, "Export Complete!", "SUCCESS");
							}
							else
							{
								MessageBox.Show(Application.Current.MainWindow, "Print Canceled! File saved in " + ndcarFileDirectory + ".", "ERROR");
							}
						}
						catch (Exception)
						{
							// Prompt with error message
							MessageBox.Show(Application.Current.MainWindow, "Export Failed!", "ERROR");
						}

						// Continue regardless to close workbook and release excel objects from memory
						excelHalfSheetWorkBook.Close(false);
						excelApp.Quit();
						Marshal.ReleaseComObject(excelHalfSheetWorkSheet);
						Marshal.ReleaseComObject(excelHalfSheetWorkBook);
						Marshal.ReleaseComObject(excelApp);
					}
					else
					{
						MessageBox.Show(Application.Current.MainWindow, "Initials needed to generate Half Sheet. Please try again.", "ERROR");
					}
				}
			}
			catch (COMException comEX)
			{
				MessageBox.Show(Application.Current.MainWindow, "Please ensure that Mircosoft Office Excel is installed on this machine.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, comEX.ToString());
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ExportSortie_HalfSheet' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to export all aircraft into their respective RAMPOD Excel files
		/// </summary>
		/// <returns>None (Void)</returns>
		private void ExportSortie_RAMPOD()
		{
			try
			{
				// Variables for CommonOpenFileDialog
				CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();
				commonOpenFileDialog.Title = "Select Folder to Export to:";
				commonOpenFileDialog.InitialDirectory = rampodFileDirectory;
				commonOpenFileDialog.IsFolderPicker = true;

				// If dialog returns a folder to use
				if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
				{
					// Set Excel file path to save files in
					string excelFilePath = commonOpenFileDialog.FileName;

					// Check if file is already open
					if (!IsFileOpen(excelFilePath + "\\" + ((DateTime)sortieDT.Rows[0][3]).ToString("d-MMM-yy").ToUpper() + ".xlsx") && 
						!IsFileOpen(excelFilePath + "\\" + ((DateTime)sortieDT.Rows[0][3]).ToString("d-MMM-yy").ToUpper() + " LOWACT.xlsx"))
					{
						// General variables
						int lowRowCount = 2;
						int highRowCount = 2;

						// Variables for Excel applications, workbooks, and worksheets
						Excel.Application excelApp = new Excel.Application();
						Excel.Workbook excelHAAWorkBook;
						Excel.Workbook excelLAAWorkBook;
						Excel.Worksheet excelHAAWorkSheet;
						Excel.Worksheet excelLAAWorkSheet;

						// Add empty values to workbooks to create them and select first worksheet
						excelHAAWorkBook = excelApp.Workbooks.Add();
						excelLAAWorkBook = excelApp.Workbooks.Add();
						excelHAAWorkSheet = (Excel.Worksheet)excelHAAWorkBook.Worksheets.get_Item(1);
						excelLAAWorkSheet = (Excel.Worksheet)excelLAAWorkBook.Worksheets.get_Item(1);

						// Setup worksheet headers
						excelLAAWorkSheet.Cells[1, 1] = excelHAAWorkSheet.Cells[1, 1] = "MISSION_ID";
						excelLAAWorkSheet.Cells[1, 2] = excelHAAWorkSheet.Cells[1, 2] = "SERNO";
						excelLAAWorkSheet.Cells[1, 3] = excelHAAWorkSheet.Cells[1, 3] = "AC_TAILNO";
						excelLAAWorkSheet.Cells[1, 4] = excelHAAWorkSheet.Cells[1, 4] = "SORTIE_DATE";
						excelLAAWorkSheet.Cells[1, 5] = excelHAAWorkSheet.Cells[1, 5] = "AC_STATION";
						excelLAAWorkSheet.Cells[1, 6] = excelHAAWorkSheet.Cells[1, 6] = "AC_TYPE";
						excelLAAWorkSheet.Cells[1, 7] = excelHAAWorkSheet.Cells[1, 7] = "CURR_UNIT";
						excelLAAWorkSheet.Cells[1, 8] = excelHAAWorkSheet.Cells[1, 8] = "ASSG_UNIT";
						excelLAAWorkSheet.Cells[1, 9] = excelHAAWorkSheet.Cells[1, 9] = "RANGE";
						excelLAAWorkSheet.Cells[1, 10] = excelHAAWorkSheet.Cells[1, 10] = "SORTIE_EFFECT";
						excelLAAWorkSheet.Cells[1, 11] = excelHAAWorkSheet.Cells[1, 11] = "IS_NON_PODDED";
						excelLAAWorkSheet.Cells[1, 12] = excelHAAWorkSheet.Cells[1, 12] = "IS_DEBRIEF";
						excelLAAWorkSheet.Cells[1, 13] = excelHAAWorkSheet.Cells[1, 13] = "IS_LIVE_MONITOR";
						excelLAAWorkSheet.Cells[1, 14] = excelHAAWorkSheet.Cells[1, 14] = "REASON";
						excelLAAWorkSheet.Cells[1, 15] = excelHAAWorkSheet.Cells[1, 15] = "REMARKS";

						// Loop through all sorties in datatable
						foreach (DataRow sortieRow in sortieDT.Rows)
						{
							// Retrieve aircraft information
							foreach (DataRow aircraftRow in sortieRow.GetChildRows(paperloadDR))
							{
								// Check if low activity
								if ((bool)aircraftRow["Aircraft_IsLowActivity"])
								{
									excelLAAWorkSheet.Cells[lowRowCount, 1] = sortieRow["Sortie_MissionID"].ToString();						// MISSION_ID          
									excelLAAWorkSheet.Cells[lowRowCount, 2] = String.Empty;													// SERNO         
									excelLAAWorkSheet.Cells[lowRowCount, 3] = aircraftRow["Aircraft_IFF"];									// AC_TAILNO                   
									excelLAAWorkSheet.Cells[lowRowCount, 4] = ((DateTime)sortieRow["Sortie_Date"]).ToString("d-MMM-yy");	// SORTIE_DATE               
									excelLAAWorkSheet.Cells[lowRowCount, 5] = String.Empty;													// AC_STATION                     
									excelLAAWorkSheet.Cells[lowRowCount, 6] = aircraftRow["Aircraft_Type"];									// AC_TYPE
									excelLAAWorkSheet.Cells[lowRowCount, 7] = "57 FW";														// CURR_UNIT
									excelLAAWorkSheet.Cells[lowRowCount, 8] = "57 FW";														// ASSG_UNIT
									excelLAAWorkSheet.Cells[lowRowCount, 9] = "NELLIS";														// RANGE
									excelLAAWorkSheet.Cells[lowRowCount, 11] = "Y";															// IS_NON_PODDED
									excelLAAWorkSheet.Cells[lowRowCount, 12] = "Y";															// IS_DEBRIEF
									excelLAAWorkSheet.Cells[lowRowCount, 13] = "Y";															// IS_LIVE_MONITOR

									// Set custom values in effect cells based on Aircraft_TrackStatus
									if (aircraftRow["Aircraft_TrackStatus"].Equals("CNX"))
									{
										excelLAAWorkSheet.Cells[lowRowCount, 10] = "OTHER";								// SORTIE_EFFECT
										excelLAAWorkSheet.Cells[lowRowCount, 14] = aircraftRow["Aircraft_TrackStatus"];	// REASON
										excelLAAWorkSheet.Cells[lowRowCount, 15] = aircraftRow["Aircraft_TrackStatus"];	// REMARKS
									}
									else
									{
										excelLAAWorkSheet.Cells[lowRowCount, 10] = "EFFECTIVE";                 // SORTIE_EFFECT
										excelLAAWorkSheet.Cells[lowRowCount, 14] = " ";                         // REASON (Have to put whitespace to avoid tripping RAMPOD error)
										excelLAAWorkSheet.Cells[lowRowCount, 15] = " ";                         // REMARKS (Have to put whitespace to avoid tripping RAMPOD error)
									}

									// Increment row count
									lowRowCount = lowRowCount + 1;
								}
								else
								{
									excelHAAWorkSheet.Cells[highRowCount, 1] = sortieRow["Sortie_MissionID"].ToString();					// MISSION_ID          
									excelHAAWorkSheet.Cells[highRowCount, 2] = "P" + aircraftRow["Aircraft_PodSerialNumber"] + "A";			// SERNO         
									excelHAAWorkSheet.Cells[highRowCount, 3] = aircraftRow["Aircraft_TailNumber"];							// AC_TAILNO                   
									excelHAAWorkSheet.Cells[highRowCount, 4] = ((DateTime)sortieRow["Sortie_Date"]).ToString("d-MMM-yy");	// SORTIE_DATE               
									excelHAAWorkSheet.Cells[highRowCount, 5] = aircraftRow["Aircraft_Station"];								// AC_STATION                     
									excelHAAWorkSheet.Cells[highRowCount, 6] = aircraftRow["Aircraft_Type"];								// AC_TYPE
									excelHAAWorkSheet.Cells[highRowCount, 7] = "57 FW";														// CURR_UNIT
									excelHAAWorkSheet.Cells[highRowCount, 8] = "57 FW";														// ASSG_UNIT
									excelHAAWorkSheet.Cells[highRowCount, 9] = "NELLIS";													// RANGE
									excelHAAWorkSheet.Cells[highRowCount, 11] = "N";														// IS_NON_PODDED
									excelHAAWorkSheet.Cells[highRowCount, 12] = "Y";														// IS_DEBRIEF
									excelHAAWorkSheet.Cells[highRowCount, 13] = "Y";														// IS_LIVE_MONITOR

									// Set custom values in effect cells based on Aircraft_TrackStatus
									if (aircraftRow["Aircraft_TrackStatus"].Equals("CNX"))
									{
										excelHAAWorkSheet.Cells[highRowCount, 10] = "OTHER";								// SORTIE_EFFECT
										excelHAAWorkSheet.Cells[highRowCount, 14] = aircraftRow["Aircraft_TrackStatus"];	// REASON
										excelHAAWorkSheet.Cells[highRowCount, 15] = aircraftRow["Aircraft_TrackStatus"];	// REMARKS
									}
									else if (aircraftRow["Aircraft_TrackStatus"].Equals("BT") || aircraftRow["Aircraft_TrackStatus"].Equals("NT"))
									{
										excelHAAWorkSheet.Cells[highRowCount, 10] = "NON-EFFECTIVE";						// SORTIE_EFFECT
										excelHAAWorkSheet.Cells[highRowCount, 14] = aircraftRow["Aircraft_TrackStatus"];	// REASON
										excelHAAWorkSheet.Cells[highRowCount, 15] = "INW";									// REMARKS
									}
									else
									{
										excelHAAWorkSheet.Cells[highRowCount, 10] = "EFFECTIVE";                 // SORTIE_EFFECT
										excelHAAWorkSheet.Cells[highRowCount, 14] = " ";                         // REASON (Have to put whitespace to avoid tripping RAMPOD error)
										excelHAAWorkSheet.Cells[highRowCount, 15] = " ";                         // REMARKS (Have to put whitespace to avoid tripping RAMPOD error)
									}

									// Increment row count
									highRowCount = highRowCount + 1;
								}
							}
						}

						// Sort "Effectiveness" column to make it easier to parse
						dynamic haaDataRange = excelHAAWorkSheet.UsedRange;
						dynamic laaDataRange = excelLAAWorkSheet.UsedRange;
						haaDataRange.sort(haaDataRange.Columns[10], Excel.XlSortOrder.xlDescending);
						laaDataRange.sort(laaDataRange.Columns[10], Excel.XlSortOrder.xlDescending);
						excelHAAWorkSheet.Columns.AutoFit();
						excelLAAWorkSheet.Columns.AutoFit();

						try
						{
							// Save workbooks
							excelHAAWorkBook.SaveAs(excelFilePath + "\\" + ((DateTime)sortieDT.Rows[0][3]).ToString("d-MMM-yy").ToUpper() + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
								Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive);
							excelLAAWorkBook.SaveAs(excelFilePath + "\\" + ((DateTime)sortieDT.Rows[0][3]).ToString("d-MMM-yy").ToUpper() + " LOWACT.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
								Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive);

							// Generate processedFileName during export to ensure that the times appended are current
							string processedFileName = Directory.GetParent(Directory.GetParent(Assembly.GetExecutingAssembly().Location).ToString()) + "\\_LMT Files\\Processed\\"
								+ ((DateTime)sortieDT.Rows[0][3]).ToString("d-MMM-yy").ToUpper() + DateTime.Now.ToString(" EXyy-MM-dd.HHmm") + ".xml";

							// Move file from Unprocessed directory to processed
							if (File.Exists(unprocessedFileName) && !File.Exists(processedFileName) && !IsFileOpen(unprocessedFileName))
							{
								File.Move(unprocessedFileName, processedFileName);
							}
							else
							{
								MessageBox.Show(Application.Current.MainWindow, "Cannot move unprocessed XML file to the processed folder.", "ERROR");
							}

							// Reset mission, player number, and reassign subheader
							missionNumber = 1;
							playerNumber = 1;
							sortieMissionIDSub_input.Text = missionNumber.ToString();
							sortieDate_input.IsEnabled = true;

							// Clear datatables
							paperloadDS.Clear();
							unprocessedFileName = Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Assembly.GetExecutingAssembly().Location))) + "\\_LMT Files\\Unprocessed\\tempFile.xml");

							// Prompt with success message
							MessageBox.Show(Application.Current.MainWindow, "Export Complete!", "SUCCESS");
						}
						catch (Exception)
						{
							// Prompt with error message
							MessageBox.Show(Application.Current.MainWindow, "Export Failed!", "ERROR");
						}

						// Continue regardless to close workbooks and release excel objects from memory
						excelHAAWorkBook.Close(false);
						excelLAAWorkBook.Close(false);
						excelApp.Quit();
						Marshal.ReleaseComObject(excelHAAWorkSheet);            
						Marshal.ReleaseComObject(excelHAAWorkBook);          
						Marshal.ReleaseComObject(excelLAAWorkSheet);  
						Marshal.ReleaseComObject(excelLAAWorkBook);  
						Marshal.ReleaseComObject(excelApp);     
					}
				}
				else
				{
					MessageBox.Show(Application.Current.MainWindow, "Could not open folder.", "ERROR");
				}

			}
			catch (COMException comEX)
			{
				MessageBox.Show(Application.Current.MainWindow, "Please ensure that Mircosoft Office Excel is installed on this machine.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, comEX.ToString());
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'ExportSortie_RAMPOD' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to generate the Mission ID julian date and populate the sortieMissionIDJulian_input field
		/// </summary>
		/// <returns>None (Void)</returns>
		private void GenerateJulianDate()
		{
			try
			{
				// Define variables to use for generation
				int tempDay = DateTime.Parse(sortieDate_input.Text).DayOfYear;
				string tempString = String.Empty;

				// Check if less than 10 or 100 to add leading zero(s)
				if (tempDay < 10)
				{
					tempString = "00" + tempDay.ToString();
				}
				else if (tempDay < 100)
				{
					tempString = "0" + tempDay.ToString();
				}
				else
				{
					tempString = tempDay.ToString();
				}

				// Assign to textbox
				sortieMissionIDJulian_input.Text = "M" + DateTime.Parse(sortieDate_input.Text).ToString("yy").Substring(1) + tempString;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'GenerateJulianDate' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to generate the Mission ID
		/// </summary>
		/// <returns>String of the Mission ID</returns>
		private string GenerateMissionID()
		{
			try
			{
				return sortieMissionIDJulian_input.Text + "-" + sortieMissionIDTime_input.Text.Replace(" ", String.Empty).ToUpper() + "-" + sortieMissionIDSub_input.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'GenerateMissionID' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return String.Empty;
			}
		}

		/// <summary>
		/// Function to check if a file is already open by another process
		/// </summary>
		/// <param name="fileName">String of the file name to check</param>
		/// <returns>Boolean if file is open or not</returns>
		private bool IsFileOpen(string fileName)
		{
			try
			{
				// Check if file exists
				if (File.Exists(fileName))
				{
					// Attempt to open file in filestream
					using (FileStream tempStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
					{
						// If opened then close stream and return false
						tempStream.Close();
						return false;
					}
				}

				// If file does not exist then return false
				return false;
			}
			catch (Exception)
			{
				// If exception is caught then file is open in another process
				MessageBox.Show(Application.Current.MainWindow, "The file you are attempting to overwrite is open in another process, please close it before continuing.", "ERROR");
				return true;
			}
		}

		/// <summary>
		/// Function to lock or unlock the sortie fields based on the specified value
		/// </summary>
		/// <param name="lockField">Boolean to denote whether the fields are locked or not</param>
		/// <returns>None (Void)</returns>
		private void LockSortieFields(bool lockField)
		{
			try
			{
				// Toggle if field is enabled based on passed boolean variable
				sortieMissionIDJulian_input.IsEnabled = !lockField;
				sortieMissionIDTime_input.IsEnabled = !lockField;
				sortieMissionIDSub_input.IsEnabled = !lockField;
				sortieStartTime_input.IsEnabled = !lockField;
				sortieEndTime_input.IsEnabled = !lockField;
				sortieProject_input.IsEnabled = !lockField;
				sortieNumCD_input.IsEnabled = !lockField;
				sortieStationM_input.IsEnabled = !lockField;
				sortieStation2_input.IsEnabled = !lockField;
				sortieStation3_input.IsEnabled = !lockField;
				sortieStation4_input.IsEnabled = !lockField;
				sortieStation5_input.IsEnabled = !lockField;
				sortieStation6_input.IsEnabled = !lockField;
				sortieStation7_input.IsEnabled = !lockField;
				sortieStation8_input.IsEnabled = !lockField;
				sortieStation9_input.IsEnabled = !lockField;
				sortieStation10_input.IsEnabled = !lockField;
				sortieDash1_input.IsEnabled = !lockField;
				sortieDash2_input.IsEnabled = !lockField;
				sortieDash3_input.IsEnabled = !lockField;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'LockSortieFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to open a file and bind the respective info to the proper datatables/datagrids
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void OpenFile_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				// Define new CommonOpenFileDialog
				CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();
				commonOpenFileDialog.Title = "Select File to Import:";
				commonOpenFileDialog.InitialDirectory = unprocessedFileDirectory;
				commonOpenFileDialog.IsFolderPicker = false;

				if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
				{
					// Save file name into global variable
					unprocessedFileName = commonOpenFileDialog.FileName;

					// Clear datatables and read opened file into respective datatables
					paperloadDS.Clear();
					paperloadDS.ReadXml(unprocessedFileName, XmlReadMode.ReadSchema);
					paperloadDS.AcceptChanges();

					// Set missionNumber
					missionNumber = 1;
					SetMaxMissionNumber();

					// Set playerNumber
					playerNumber = 1;
					SetMaxPlayerNumber();

					// Function calls to clear all fields
					AircraftClearFields();
					SortieClearFields();

					// Check to see if sortie fields need to be repopulated
					if (UpdateRowCount() != 0)
					{
						RepopulateSortieFields();
					}
				}
				else
				{
					MessageBox.Show(Application.Current.MainWindow, "Could not open file.", "ERROR");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'OpenFile_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to rebuild the pod serial dropdown list
		/// </summary>
		/// <returns>None (Void)</returns>
		private void RebuildPodSerialList()
		{
			try
			{
				// Create list to use for Aircraft Pod Serials and add default values
				List<string> aircraftPodSerials = new List<string>();
				aircraftPodSerials.AddRange("50001,50002,50014,50018,50023,50024,50029,50030,50031,50033,50066,50067,50073,50075,50079,50080,50082,50086,50087,50092,50093,50094,50096,50097,50103,50104,50106,50131,50132,50141,50143,50144,50203,50211,50213,50214,50215,50216,50217,50218,50219,50220,50221,50222,50223,50307,50308,50329,50330,50331,50332,50333,50334,50335,50336,50337,50338,50339,50340,50341,50342,50380,50383,50384,50505,50514,50515,50517,50519,50522,50677,50679,50683,50684,50693,50694,50695,50696,50697,50698,50699,50700,50701,50702,50703,50704,50705,50706,50707,50708,50709,50710,50726,50727,50728,50729,50730,50731,50732,50733,50741,50750,50759,51489,51490,51491,51492,51493,51494,51495,51496,51497,51498,51499,51500,51501,51502,51503,51504,51505,51506,51507,51508,51509,51510,51511,51512,51513,51514,51515,51516,51517,51518,51519,51520,51521,51522,51523,51524,51525,51526,51527,51528,51529,51530,51531,51532,51533,51534,51535".Split(',').ToList());

				// Clear existing pod serials and add ones from list
				aircraftPodSerial_input.Items.Clear();
				foreach (string aircraftPodSerial in aircraftPodSerials)
				{
					aircraftPodSerial_input.Items.Add(aircraftPodSerial);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'RebuildPodSerialList' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function that returns the respective stations the sortie was recorded on
		/// </summary>
		/// <returns>String of the stations the sortie was recorded on</returns>
		private string RecordedStations()
		{
			try
			{
				// Create temp string to return.
				string tempStation = String.Empty;
				string tempDash = String.Empty;

				// Check to see which radiobutton is selected
				if ((bool)sortieDash1_input.IsChecked)
				{
					tempDash = "(-1)";
				}
				else if ((bool)sortieDash2_input.IsChecked)
				{
					tempDash = "(-2)";
				}
				else
				{
					tempDash = "(-3)";
				}

				// Check to see which checkboxes are checked
				if ((bool)sortieStationM_input.IsChecked)
				{
					tempStation = "M,";
				}
				if ((bool)sortieStation2_input.IsChecked)
				{
					tempStation = tempStation + "2,";
				}
				if ((bool)sortieStation3_input.IsChecked)
				{
					tempStation = tempStation + "3,";
				}
				if ((bool)sortieStation4_input.IsChecked)
				{
					tempStation = tempStation + "4,";
				}
				if ((bool)sortieStation5_input.IsChecked)
				{
					tempStation = tempStation + "5,";
				}
				if ((bool)sortieStation6_input.IsChecked)
				{
					tempStation = tempStation + "6,";
				}
				if ((bool)sortieStation7_input.IsChecked)
				{
					tempStation = tempStation + "7,";
				}
				if ((bool)sortieStation8_input.IsChecked)
				{
					tempStation = tempStation + "8,";
				}
				if ((bool)sortieStation9_input.IsChecked)
				{
					tempStation = tempStation + "9,";
				}
				if ((bool)sortieStation10_input.IsChecked)
				{
					tempStation = tempStation + "10,";
				}

				// If no station is selected then return an empty string 
				if (String.IsNullOrEmpty(tempStation))
				{
					return String.Empty;
				}

				// Replace last comma with a colon and add mission header dash
				return "CS," + tempStation.Remove(tempStation.Length - 1) + ":" + tempDash;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'EmptyFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return String.Empty;
			}
		}

		/// <summary>
		/// Function to remove a pod serial number from the respective dropdown list
		/// </summary>
		/// <param name="serial">String of the pod serial to remove from the dropdown list</param>
		/// <returns>None (Void)</returns>
		private void RemovePodSerialNumber(string serial)
		{
			try
			{
				// Check if LAA
				if (!serial.Equals("N/A"))
				{
					foreach (string tempItem in aircraftPodSerial_input.Items)
					{
						// If item from dropdown matches passed parameter then remove and break out of loop
						if (tempItem.Equals(serial))
						{
							aircraftPodSerial_input.Items.Remove(tempItem);
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'RemovePodSerialNumber' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to repopulate the sortie fields if there is an unsubmitted sortie
		/// </summary>
		/// <returns>None (Void)</returns>
		private void RepopulateSortieFields()
		{
			try
			{
				foreach (DataRow sortieRow in sortieDT.Rows)
				{
					if (!(bool)sortieRow["Sortie_IsMissionSubmitted"])
					{
						// Set sortie info (even though it loops through each child row, there should only be one, break just in case)
						foreach (DataRow aircraftRow in sortieRow.GetChildRows(paperloadDR))
						{
							// Remove serial number from dropdown list
							RemovePodSerialNumber(aircraftRow["Aircraft_PodSerialNumber"].ToString());
						}

						// Create temp variables for DateTimes, RecordingStations, and NumofCDs to convert
						string[] tempID = sortieRow["Sortie_MissionID"].ToString().Split("-");
						string tempRecordDash = sortieRow["Sortie_RecordingStations"].ToString();
						string[] tempRecords = tempRecordDash.Split(",");

						// Resign values to sortie fields
						sortieMissionIDJulian_input.Text = tempID[0];
						sortieMissionIDTime_input.Text = tempID[1];
						sortieMissionIDSub_input.Text = tempID[2];
						sortieDate_input.SelectedDate = (DateTime)sortieRow["Sortie_Date"];
						sortieStartTime_input.Text = ((DateTime)sortieRow["Sortie_StartRangeTime"]).ToString("HHmm");
						sortieEndTime_input.Text = ((DateTime)sortieRow["Sortie_EndRangeTime"]).ToString("HHmm");
						sortieProject_input.Text = sortieRow["Sortie_ProjectNumber"].ToString();
						sortieNumCD_input.Text = ((int)sortieRow["Sortie_NumOfCDs"]).ToString();

						// Reselect recorded station radiobutton
						if (tempRecordDash.Substring(tempRecordDash.Length - 2, 1) == "1")
						{
							sortieDash1_input.IsChecked = true;
						}
						else if (tempRecordDash.Substring(tempRecordDash.Length - 2, 1) == "2")
						{
							sortieDash2_input.IsChecked = true;
						}
						else if (tempRecordDash.Substring(tempRecordDash.Length - 2, 1) == "3")
						{
							sortieDash3_input.IsChecked = true;
						}

						foreach (string tempRecord in tempRecords)
						{
							// Repopulate recorded station fields
							if (tempRecord.Substring(0, 1) == "M")
							{
								sortieStationM_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "2")
							{
								sortieStation2_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "3")
							{
								sortieStation3_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "4")
							{
								sortieStation4_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "5")
							{
								sortieStation5_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "6")
							{
								sortieStation6_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "7")
							{
								sortieStation7_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "8")
							{
								sortieStation8_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "9")
							{
								sortieStation9_input.IsChecked = true;
							}
							if (tempRecord.Substring(0, 1) == "1")
							{
								sortieStation10_input.IsChecked = true;
							}
						}

						// Break out of sortie loop since only a single sortie needs to be repopulated
						break;
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'RepopulateSortieFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the selected aircraft via the respective callsign, broken out from text changed event call function
		/// </summary>
		/// <returns>None (Void)</returns>
		private void SelectAircraftFromCallsign()
		{
			try
			{
				if (aircraftUnit_input.Text == "422 TES")
				{
					// Set up temp empty string to check by
					string tempAircraft = String.Empty;

					// Check for inputed value in "Unit" textbox
					if (aircraftCallsign_input.Text == "STRIKE" || aircraftCallsign_input.Text == "EAGLE")
					{
						tempAircraft = "F-15";
					}
					else if (aircraftCallsign_input.Text == "RAPTOR")
					{
						tempAircraft = "F-22A";
					}
					else if (aircraftCallsign_input.Text == "BOLT")
					{
						tempAircraft = "F-35A";
					}
					else if (aircraftCallsign_input.Text == "VIPER" || aircraftCallsign_input.Text == "VENOM")
					{
						tempAircraft = "F-16";
					}
					else if (aircraftCallsign_input.Text == "BOAR")
					{
						tempAircraft = "A-10";
					}

					// If no correct entry then do not select any entry and leave box blank
					if (tempAircraft == String.Empty)
					{
						aircraftType_input.SelectedIndex = -1;
					}
					else
					{
						// Check each of the available aircraft options and select the correct entry
						foreach (string item in aircraftType_input.Items)
						{
							if (item.Equals(tempAircraft))
							{
								aircraftType_input.SelectedItem = item;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SelectAircraftFromCallsign' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the selected aircraft via the respective unit, broken out from text changed event call function
		/// </summary>
		/// <returns>None (Void)</returns>
		private void SelectAircraftFromUnit()
		{
			try
			{
				// Set up temp empty string to check by
				string tempAircraft = String.Empty;

				// Check for inputed value in "Unit" textbox
				if (aircraftUnit_input.Text == "64 AGRS" || aircraftUnit_input.Text == "16 WPS" || aircraftUnit_input.Text == "TOP ACES")
				{
					tempAircraft = "F-16";
				}
				else if (aircraftUnit_input.Text == "17 WPS")
				{
					tempAircraft = "F-15";
				}
				else if (aircraftUnit_input.Text == "6 WPS" || aircraftUnit_input.Text == "65 AGRS")
				{
					tempAircraft = "F-35A";
				}
				else if (aircraftUnit_input.Text == "66 WPS")
				{
					tempAircraft = "A-10";
				}
				else if (aircraftUnit_input.Text == "433 WPS")
				{
					tempAircraft = "F-22A";
				}
				else if (aircraftUnit_input.Text == "26 WPS")
				{
					tempAircraft = "MQ-9";
				}

				// If no correct entry then do not select any entry and leave box blank
				if (tempAircraft == String.Empty)
				{
					aircraftType_input.SelectedIndex = -1;
				}
				else
				{
					// Check each of the available aircraft options and select the correct entry
					foreach (string item in aircraftType_input.Items)
					{
						if (item.Equals(tempAircraft))
						{
							aircraftType_input.SelectedItem = item;
						}
					}
				}

				// Call function to change available units
				ChangeCallsignOptions();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SelectAircraftFromUnit' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the selected aircraft via the respective callsign (Mainly for 422TES)
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any text changed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SelectAircraftFromCallsign_TextChanged(object sender, TextChangedEventArgs e)
		{
			try
			{
				SelectAircraftFromCallsign();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SelectAircraftFromCallsign_TextChanged' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to change the selected aircraft via the respective unit
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any text changed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SelectAircraftFromUnit_TextChanged(object sender, TextChangedEventArgs e)
		{
			try
			{
				SelectAircraftFromUnit();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SelectAircraftFromUnit_TextChanged' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to set the Mission ID subheader
		/// </summary>
		private void SetMissionIDSub()
		{
			try
			{
				sortieMissionIDSub_input.Text = missionNumber.ToString();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SetMissionIDSub' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to get and set the max mission ID number
		/// </summary>
		/// <returns>None (Void)</returns>
		private void SetMaxMissionNumber()
		{
			try
			{
				// Check all rows in sortie datatable
				foreach (DataRow sortieRow in sortieDT.Rows)
				{
					// Check to make sure row hasn't been submitted yet
					if (!(bool)sortieRow["Sortie_IsMissionSubmitted"])
					{
						missionNumber = (int)sortieRow["Sortie_MissionNumber"];
						break;
					}
					// If missionNumber is currently less than the parsed missionNumber
					else if (missionNumber < (int)sortieRow["Sortie_MissionNumber"])
					{
						missionNumber = (int)sortieRow["Sortie_MissionNumber"];
					}
				}

				// Check to see if missionNumber needs to be incremented
				if (UpdateRowCount() == 0 && sortieDT.Rows.Count != 0)
				{
					missionNumber = missionNumber + 1;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SetMaxMissionNumber' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to get and set the max player number
		/// </summary>
		/// <returns>None (Void)</returns>
		private void SetMaxPlayerNumber()
		{
			try
			{
				// Check all rows in datatable
				foreach (DataRow aircraftRow in aircraftDT.Rows)
				{
					// If playerNumber is currently less than the parsed playerNumber
					if (playerNumber < (int)aircraftRow["Aircraft_PlayerNumber"])
					{
						playerNumber = (int)aircraftRow["Aircraft_PlayerNumber"];
					}
				}

				// Check to see if playerNumber needs to be incremented
				if (aircraftDT.Rows.Count != 0)
				{
					playerNumber = playerNumber + 1;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SetMaxPlayerNumber' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to clear all needed fields after submitting a sortie
		/// </summary>
		/// <returns>None (Void)</returns>
		private void SortieClearFields()
		{
			try
			{
				// Empty all fields
				sortieMissionIDTime_input.Text = String.Empty;
				SetMissionIDSub();
				sortieStartTime_input.Text = String.Empty;
				sortieEndTime_input.Text = String.Empty;
				sortieProject_input.Text = String.Empty;
				sortieNumCD_input.Text = String.Empty;
				sortieStationM_input.IsChecked = false;
				sortieStation2_input.IsChecked = false;
				sortieStation3_input.IsChecked = false;
				sortieStation4_input.IsChecked = false;
				sortieStation5_input.IsChecked = false;
				sortieStation6_input.IsChecked = false;
				sortieStation7_input.IsChecked = false;
				sortieStation8_input.IsChecked = false;
				sortieStation9_input.IsChecked = false;
				sortieStation10_input.IsChecked = false;
				aircraftLowAct_input.IsChecked = false;
				aircraftUnit_input.Text = String.Empty;
				aircraftCallsign_input.Text = String.Empty;
				aircraftType_input.Text = String.Empty;
				aircraftStation_input.Text = String.Empty;

				// Change inputs since LowAct box was reset
				AircraftInputModeChange();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieClearFields' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to load all players into aircraft datagrid when sortie title is double clicked in the sortie datagrid
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any mouse button event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			try
			{
				if (sortieDG.SelectedItem != null && UpdateRowCount() != 0)
				{
					MessageBox.Show(Application.Current.MainWindow, "Please submit the current sortie or delete all of the players before loading a new mission.", "ERROR");
				}
				else if (sortieDG.SelectedItem != null && aircraftAddSortie_button.IsEnabled)
				{
					// Get currently selected row
					DataRowView selectedRow = sortieDG.SelectedItem as DataRowView;
					DataRow[] sortieRow = sortieDT.Select("Sortie_MissionNumber = " + selectedRow.Row[0].ToString());

					// Set mission status, current mission number, and change statuses
					sortieRow[0]["Sortie_IsMissionSubmitted"] = false;
					missionNumber = (int)sortieRow[0]["Sortie_MissionNumber"];
					ChangePlayersSubmitted(missionNumber, false);

					// Clear aircraft and sortie fields and repopulate sortie fields
					AircraftClearFields();
					SortieClearFields();
					RepopulateSortieFields();
					UpdateRowCount();

					// Write changes to the XML file
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDataGrid_MouseDoubleClick' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to open the secondary auditorium window
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDataGridContextAddAud_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				AuditoriumWindow auditoriumWindow = new AuditoriumWindow();
				auditoriumWindow.Owner = this;
				auditoriumWindow.ShowDialog();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDataGridContextAddAud_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to add notes to the selected sortie
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDataGridContextAddNotes_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (sortieDG.SelectedItem != null)
				{
					// Get currently selected row
					DataRowView selectedRow = sortieDG.SelectedItem as DataRowView;
					DataRow[] sortieRow = sortieDT.Select("Sortie_MissionNumber = " + selectedRow.Row[0].ToString());

					// Save returned text to sortie datatable and write changes to XML file
					sortieRow[0]["Sortie_Auditoriums"] = Prompt.ShowDialog("Please enter your notes to add to the Half Sheet below.", "ADD NOTES", sortieRow[0]["Sortie_Auditoriums"].ToString(), 848, 185, 763, 52);
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDataGridContextAddNotes_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to delete the sortie and its respective aircraft when the "Delete Sortie" option is selected in the sortie datagrid context menu
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDataGridContextDelete_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (MessageBox.Show(Application.Current.MainWindow, "Are you sure you want to delete this sortie?", "WARNING", MessageBoxButton.YesNo) == MessageBoxResult.Yes && sortieDG.SelectedItem != null)
				{
					// Get currently selected row
					DataRowView selectedRow = sortieDG.SelectedItem as DataRowView;
					DataRow[] sortieRow = sortieDT.Select("Sortie_MissionNumber = " + selectedRow.Row[0].ToString());

					// Delete all child rows of selected sortie
					foreach (DataRow aircraftRow in sortieRow[0].GetChildRows(paperloadDR))
					{
						aircraftRow.Delete();
					}

					// Delete the correct row from the sortie datatable, accept changes in both datatables, and write to XML file
					sortieRow[0].Delete();
					paperloadDS.AcceptChanges();
					paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDataGridContextDelete_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to call the function ExportSortie_HalfSheet when "Generate Half Sheet" is selected in the sortie context menu
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDataGridContextGenerate_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (sortieDG.SelectedItem != null)
				{
					ExportSortie_HalfSheet();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDataGridContextGenerate_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to set the text in sortieMissionIDJulian_input when the sortieDate_input textbox is loaded
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDate_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				GenerateJulianDate();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDate_Loaded' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to set the text in sortieMissionIDJulian_input when the selection in sortieDate_input is changed
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any selection changed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
		{
			try
			{
				GenerateJulianDate();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieDate_SelectedDateChanged' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function for the sortieRAMPODExport_button to call the ExportSortie_RAMPOD function
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieRAMPODExport_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				// Check if all aircraft have been submitted
				if (UpdateRowCount() != 0)
				{
					MessageBox.Show(Application.Current.MainWindow, "Please submit all aircraft before exporting sorties.", "ERROR");
				}
				// Check if datatable is empty
				else if (aircraftDT.Rows.Count == 0)
				{
					MessageBox.Show(Application.Current.MainWindow, "No sorties to export.", "ERROR");
				}
				else
				{
					ExportSortie_RAMPOD();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieExport_Click' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to assign Mission ID Subheader on textbox load
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieMissionIDSub_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				SetMissionIDSub();
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieMissionIDSub_Loaded' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to allow modification of sortie data and changes data in XML file
		/// </summary>
		/// <param name="sender">Object that sent the action</param>
		/// <param name="e">Any routed event args sent by sender</param>
		/// <returns>None (Void)</returns>
		private void SortieModify_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if ((bool)sortieModify_input.IsChecked)
				{
					// Unlock fields and disable buttons
					MessageBox.Show(Application.Current.MainWindow, "Please uncheck the box after making your changes to submit the modified data before continuing.", "ALERT");
					LockSortieFields(false);
					DisableButtons(true, true);
					sortieModify_input.Content = "Uncheck to Submit Modified Sortie Data";

					// If sortie datable is empty then re-enable the sortieDate_input datepicker
					if (sortieDT.Rows.Count == 0)
					{
						sortieDate_input.IsEnabled = true;
					}
				}
				else
				{
					string tempFields = EmptyFields(true);

					// Check to see if any fields are empty
					if (!String.IsNullOrEmpty(tempFields))
					{
						MessageBox.Show(Application.Current.MainWindow, "Please fill out or modify the following fields:" + tempFields, "ERROR");
						sortieModify_input.IsChecked = true;
					}
					else
					{
						// Re-lock fields and re-enable buttons
						LockSortieFields(true);
						DisableButtons(false, false);
						sortieDate_input.IsEnabled = false;
						sortieModify_input.Content = "Check to Modify Sortie Data";

						// Find the row in the datatable with the specified missionNumber and change modified sortie data
						foreach (DataRow sortieRow in sortieDT.Rows)
						{
							if ((int)sortieRow["Sortie_MissionNumber"] == missionNumber)
							{
								sortieRow["Sortie_MissionID"] = GenerateMissionID();
								sortieRow["Sortie_Date"] = sortieDate_input.Text;
								sortieRow["Sortie_StartRangeTime"] = ConvertRangeTimes(true);
								sortieRow["Sortie_EndRangeTime"] = ConvertRangeTimes(false);
								sortieRow["Sortie_ProjectNumber"] = sortieProject_input.Text.Replace(" ", String.Empty).ToUpper();
								sortieRow["Sortie_NumOfCDs"] = int.Parse(sortieNumCD_input.Text.Replace(" ", String.Empty));
								sortieRow["Sortie_RecordingStations"] = RecordedStations();
								break;
							}
						}

						// Write changes to datatable and to XML file
						paperloadDS.AcceptChanges();
						paperloadDS.WriteXml(unprocessedFileName, XmlWriteMode.WriteSchema);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'SortieModify_Checked' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
			}
		}

		/// <summary>
		/// Function to update the row count in the status bar
		/// </summary>
		/// <returns>None (Void)</returns>
		private int UpdateRowCount()
		{
			try
			{
				// Set up tempCount variable
				int tempLow = 0;
				int tempHigh = 0;

				// Accept changes before parsing datatable
				paperloadDS.AcceptChanges();

				foreach (DataRow sortieRow in sortieDT.Rows)
				{
					if ((int)sortieRow["Sortie_MissionNumber"] == missionNumber && !(bool)sortieRow["Sortie_IsMissionSubmitted"])
					{
						foreach (DataRow aircraftRow in sortieRow.GetChildRows(paperloadDR))
						{
							// Check to make sure row isn't marked to be deleted
							if (!aircraftRow.RowState.Equals(DataRowState.Deleted))
							{
								if ((bool)aircraftRow["Aircraft_IsLowActivity"])
								{
									tempLow = tempLow + 1;
								}
								else
								{
									tempHigh = tempHigh + 1;
								}
							}
						}

						break;
					}
				}

				// If aircraftDT is not empty then set assigned date from first row in sortieDT and disable sortieDate_input datepicker
				if (aircraftDT.Rows.Count != 0)
				{
					sortieDate_input.SelectedDate = (DateTime)sortieDT.Rows[0][3];
					sortieDate_input.IsEnabled = false;
				}
				else
				{
					sortieDate_input.IsEnabled = true;
				}

				// Check to see if sortie fields need to be unlocked or locked
				if (tempLow + tempHigh == 0)
				{
					LockSortieFields(false);
					sortieModify_input.Visibility = Visibility.Hidden;
				}
				else
				{
					LockSortieFields(true);
					sortieModify_input.Visibility = Visibility.Visible;
				}

				// Set row labels and return row count
				rowCountLAA_label.Content = tempLow;
				rowCountHAA_label.Content = tempHigh;
				rowCount_label.Content = tempHigh + tempLow;
				return (int)rowCount_label.Content;
			}
			catch (Exception ex)
			{
				MessageBox.Show(Application.Current.MainWindow, "There was a failure in the 'UpdateRowCount' function. Restart the application and try again or contact your system administrator if the problem persists.", "ERROR");
				MessageBox.Show(Application.Current.MainWindow, ex.ToString());
				return 0;
			}
		}
	}
}