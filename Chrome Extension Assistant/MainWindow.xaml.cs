/**
 * Copyright 2020 Sakkyoi.
 * License: LGPL-3.0
 * Author: Sakkyoi[https://github.com/sakkyoi]
 * Github: https://github.com/sakkyoi/Chrome-Extension-Assistant
 * Redme: https://github.com/sakkyoi/Chrome-Extension-Assistant/blob/master/README.md
 */
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using MaterialDesignThemes.Wpf;
using System.IO;
using System.Threading;

namespace Chrome_Extension_Assistant
{
    public partial class MainWindow : Window
    {

        // Program need UAC elevation. Detail in app.manifest line 7.

        public MainWindow()
        {
            InitializeComponent();
            LoadExtensionIdList();
        }

        // Load extension id from registry(HKLM\Software\Policies\Google\Chrome\ExtensionInstallWhitelist). If is null, call IfNeedCreateExtensionWhiteList() to check is exist or not
        public void LoadExtensionIdList()
        {
            ExtensionId.Items.Clear(); // Clear ExtensionId ViewList
            RegistryKey RegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist", RegistryKeyPermissionCheck.ReadWriteSubTree); // Path to Chrome extension whitelist

            // If registry key not null, find them and add to viewlist
            if (RegKey != null)
            {
                int CountOfId = RegKey.ValueCount; // Count how many id whitelist have
                for (int i = 1; i <= CountOfId; i++)
                {
                    string RegValue = (string)RegKey.GetValue(i.ToString()); // Get id
                    ExtensionId.Items.Add(RegValue); // Add to viewlist
                }
                ExtensionIdHeader.Header = "Extension ID(" + CountOfId.ToString() + ")"; // Set viewlist columns name to "Extension ID($Numbers of ID)"
            }
            else
            {
                IfNeedCreateExtensionWhiteList(); 
                ExtensionIdHeader.Header = "Extension ID(0)"; // Set viewlist columns name to "Extension ID($Numbers of ID=0)"
            }
        }

        // IfNeedCreateExtensionWhiteList() Use to check registry key(HKLM\Software\Policies\Google\Chrome\ExtensionInstallWhitelist) is exist or not. If is not exist, create it
        public void IfNeedCreateExtensionWhiteList()
        {
            // Step step to HKLM\Software\Policies\Google\Chrome\ExtensionInstallWhitelist check if it is exist. If not exist, create it
            RegistryKey CreateRegKey = Registry.LocalMachine.OpenSubKey("Software", RegistryKeyPermissionCheck.ReadWriteSubTree);
            CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies");
            if (CreateRegKey == null)
            {
                Registry.LocalMachine.CreateSubKey("Software\\Policies");
                CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies", RegistryKeyPermissionCheck.ReadWriteSubTree);
            }
            CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google", RegistryKeyPermissionCheck.ReadWriteSubTree);
            if (CreateRegKey == null)
            {
                Registry.LocalMachine.CreateSubKey("Software\\Policies\\Google");
                CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google", RegistryKeyPermissionCheck.ReadWriteSubTree);
            }
            CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome", RegistryKeyPermissionCheck.ReadWriteSubTree);
            if (CreateRegKey == null)
            {
                Registry.LocalMachine.CreateSubKey("Software\\Policies\\Google\\Chrome");
                CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome", RegistryKeyPermissionCheck.ReadWriteSubTree);
            }
            CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist", RegistryKeyPermissionCheck.ReadWriteSubTree);
            if (CreateRegKey == null)
            {
                Registry.LocalMachine.CreateSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist");
                CreateRegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist", RegistryKeyPermissionCheck.ReadWriteSubTree);
            }
        }

        // Click add button in add dialog. 
        public void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string ExtensionId = AddExtensionId.Text; // Get Id from extension id input field

            // Check is Id user input illegal(ID must have 32 english letter)
            if (ExtensionId == "" || ExtensionId.Length != 32)
            {
                MessageDialogShow("error", "Invalid Extension ID"); // Show error message
            }
            else
            {
                RegistryKey RegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist", RegistryKeyPermissionCheck.ReadWriteSubTree);
                int IdIndex = RegKey.ValueCount + 1; // Found what name must this id be
                RegKey.SetValue(IdIndex.ToString(), AddExtensionId.Text); // Add id to whitelist
                ForceGPUpdate(); // GPUPDATE /force
                MessageDialogShow("info", "Add Success! ID:\n" + AddExtensionId.Text); // Show success message
                LoadExtensionIdList(); // Reload extension id list
                AddDialog.IsOpen = false; // Close add dialog
            }
        }

        // Double click extension ID view list item to delete id from whitelist
        private void ExtensionId_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Check if select an item before double click
            if (ExtensionId.SelectedItem != null)
            {
                DeleteDialog.IsOpen = true; // Open dialog to ask user is sure to delete id from whitelist 
                DeleteDialogContent.Text = "Are you sure to delete extension ID:\n" + ExtensionId.SelectedItem.ToString() + "?"; // Dialog message: "Are you sure to delete extension ID:\n$SelectedItem?
            }
        }

        // Cancel selected item while click
        private void ExtensionIdSelectedItem_Cancel(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Check if asking delete
            if (DeleteDialog.IsOpen != true)
            {
                ExtensionId.SelectedItem = null;
            }
        }

        // Click sure button in delete dialog
        public void SureDeleteButton_Click(object sender, RoutedEventArgs e)
        {
            RegistryKey RegKey = Registry.LocalMachine.OpenSubKey("Software\\Policies\\Google\\Chrome\\ExtensionInstallWhitelist", RegistryKeyPermissionCheck.ReadWriteSubTree);
            int CountOfId = RegKey.ValueCount; // Count how many id whitelist have

            // Step step to check which index of id is id to delete
            for (int i = 1; i <= CountOfId; i++)
            {
                string RegValue = (string)RegKey.GetValue(i.ToString()); // Get id from whitelist to check

                // Found id, set id to next value
                if (RegValue == ExtensionId.SelectedItem.ToString())
                {

                    // Set id to next value
                    for (int x = i; x <= CountOfId; x++)
                    {
                        if (x != CountOfId)
                        {
                            RegKey.SetValue(x.ToString(), RegKey.GetValue((x + 1).ToString())); // Also have next id
                        }
                        else
                        {
                            RegKey.DeleteValue(x.ToString()); // Last id
                        }
                    }
                    break; // break loop
                }
            }
            DeleteDialog.IsOpen = false; // Close delete dialog
            DeleteDialogContent.Text = ""; // Clear dialog content
            ExtensionId.SelectedItem = null; // cancel selected
            ForceGPUpdate(); // gpupdate /force
            LoadExtensionIdList(); // Reload extension id list
        }

        // Click cancel button in delete dialog
        private void CancelDeleteButton_Click(object sender, RoutedEventArgs e)
        {
            ExtensionId.SelectedItem = null; // cancel selected
        }

        // gpupdate /force
        public void ForceGPUpdate()
        {
            FileInfo execFile = new FileInfo("gpupdate.exe"); // Get gpupdate.exe from system
            Process proc = new Process(); // new process to run gpupdate.exe
            proc.StartInfo.CreateNoWindow = true; // run in background
            proc.StartInfo.FileName = execFile.Name; // set execute programe to gpupdate.exe
            proc.StartInfo.Arguments = "/force"; // Set parameter /force
            proc.Start(); // Execute

            // freeze program until finished
            while (!proc.HasExited)
            {
                Thread.Sleep(100);
            }
        }

        // Message dialog
        public void MessageDialogShow(string type, string message)
        {
            switch (type)
            {
                case "info":
                    MessageDialogIcon.Kind = PackIconKind.InfoCircle;
                    MessageDialogIcon.Foreground = System.Windows.Media.Brushes.LightSlateGray;
                    break;
                case "warning":
                    MessageDialogIcon.Kind = PackIconKind.Alert;
                    MessageDialogIcon.Foreground = System.Windows.Media.Brushes.Gold;
                    break;
                case "error":
                    MessageDialogIcon.Kind = PackIconKind.Alert;
                    MessageDialogIcon.Foreground = System.Windows.Media.Brushes.Firebrick;
                    break;
            }
            MessageDialogContent.Text = message;
            MessageDialog.IsOpen = true;
        }

        // Click cancel in message dialog
        public void MessageDialog_Cancel(object sender, MaterialDesignThemes.Wpf.DialogClosingEventArgs eventArgs)
        {
            MessageDialogContent.Text = ""; // Clear message dialog content
            MessageDialogIcon.Foreground = System.Windows.Media.Brushes.LightSlateGray; // Set icon color to gray
        }

    }

}
