#This script will create a windows form with buttons that will apply different groups and OUs to the selected machines


#ASSEMBLIES
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Collections

[System.Windows.Forms.Application]::EnableVisualStyles()

    

# Create Icon Extractor Assembly
# Code found online that extracts icons from dll files
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
    [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing


#GLOBAL VARIABLES
    
    #Icons
    $form_ico = [System.IconExtractor]::Extract("C:\Windows\System32\imageres.dll", -5374, $true) 
    $search_ico = [System.IconExtractor]::Extract("C:\Windows\System32\imageres.dll", -18, $false)
    $queue_ico = [System.IconExtractor]::Extract("C:\Windows\System32\wmploc.dll", 11, $false)
    $clear_ico = [System.IconExtractor]::Extract("C:\Windows\System32\shell32.dll", 131, $false)

    #Button Sound
    $Global:bttnsound = New-Object System.Media.SoundPlayer
    $Global:bttnsound.SoundLocation = 'C:\Windows\Media\Windows Unlock.wav'

    #AD Functions  
    Function Global:Get_Devices() #returns array of computer objects in the devicebox
    {
        $global:dev_group = New-Object System.Collections.ArrayList
        For($i=0; $i -lt $Devicebox.Rows.Count; $i+=1)
        {
            $device = $DeviceBox.Rows[$i].Cells[0].value
            $getpc = Get-ADComputer -Filter 'Name -eq $device'
            $null = $global:dev_group.Add($getpc)
        }
        $dev_group 
    } 
    Function Global:Add_ADGroup()
    {        
        param($g, $m)
        Get-ADGroup -Filter 'Name -eq $g' | Add-ADGroupMember -Members $m
    }     
    
    Function Global:moveOU()
    {        
        param($org, $id)
        $ou = Get-ADOrganizationalUnit -Filter 'Name -eq $org'
        Move-ADObject -Identity $id -TargetPath $ou
    }

	#$Global:ou_search= Get-ADOrganizationalUnit -Filter 'Name -like $global:x'
        
    #Click functions
    Function global:move_selected ()
    {
        Switch($radchk)
        {
            'Device Name' {$DataBox.SelectedCells | % {$DeviceBox.Rows.Add($_.value)}} #For each selected row add it to the corresponding datagrid
            'Organizational Unit' {$DataBox.SelectedCells | % {$OUBox.Rows.Add($_.value)}}
            'Group Attribute' {$DataBox.SelectedCells | % {$GroupBox.Rows.Add($_.value)}}
        }
    }

    Function global:moveOU_click () #adds queued devices to a OU
    {
        $devices = Global:Get_Devices
        $orgunit = $OUBox.Rows[0].Cells[0].value 
        $devices | % {Global:moveOU -org $orgunit -id $_}
    }

    Function global:addgrp_click() #adds queued groups to an array of devices
    {
        $mem = Global:Get_Devices
        For($i=0; $i -lt $GroupBox.Rows.Count; $i+=1)
        {
            $grp = $GroupBox.Rows[$i].Cells[0].value
            Global:Add_ADGroup -g $grp -m $mem
        } 
    }#end

    Function Global:ClearView ()
    {
        Switch($this.name)
        {
            'button_Clear' {$ADText.Clear()}
            'DeviceBox' {$DeviceBox.Rows.Clear()}
            'OUBox' {$OUBox.Rows.Clear()}
            'GroupBox' {$GroupBox.Rows.Clear()}
        }
    }

    $global:x = ''
    $global:radchk = ''
    Function radioevent() #onclick radio function
    {
        Switch($this.name)
            {
                'OURadio' {$global:radchk = 'Organizational Unit'}
                'GARadio' {$global:radchk = 'Group Attribute'}
                'DNRadio' {$global:radchk = 'Device Name'}
            }
        $rad_grpbox.text = $radchk + " Search"
        $Databox.Rows.Clear()
        #[System.Windows.Forms.MessageBox]::Show($global:radchk, "Test")
        
        

    }
    #Search function -- Searchs for an AD Object based on the checked radio and user input ($x) that is the 
    Function global:searchfnc 
    {
        Switch($radchk) 
            {
                'Organizational Unit' {Get-ADOrganizationalUnit -Filter 'Name -like $x'}#get OUs
                'Group Attribute' {Get-ADGroup -Filter 'Name -like $x'}#get Groups
                'Device Name' {Get-ADComputer -Filter 'Name -like $x'}#get computer
            }
    }
    
    Function global:search_results #Sends search results to queue depending on the type of object returned (computers, groups or OUs)
    {
        param ($terms)
        foreach($r in $terms)
        {
            Switch($r.OBjectClass)
            {
                computer {$DataBox.Rows.Add($r.Name)}
                group {$Databox.Rows.Add($r.SamAccountName)} #there is no name obj class for groups and the SAM name of groups doesn't end in $ like computers
                organizationalUnit {$Databox.Rows.Add($r.Name)}
            }

        }
    }
    Function search_click()#captures text, adds wildcards and input to search function
    {
            #capture input and add wildcard
        if($ADText.Text)
        {   
            $x = '*' + $ADText.Text + '*'
            $search_terms = searchfnc
            $Databox.Rows.Clear()
            search_results -Terms $search_terms			
        }
        else
        {   #unfocus textbox and display watermark
            [System.Windows.Forms.MessageBox]::Show("Enter Text" , "Error")
        }
    }#end

    #Tooltips
    $search_tip = New-Object System.Windows.Forms.ToolTip
    $ShowTip =
    {
        Switch($this.name) 
        {
            'OURadio' {$tip = 'Find an OU'}
            'GARadio' {$tip = 'Find an AD Group'}
            'DNRadio' {$tip = 'Find a Device'}
            'button_Search' {$tip = 'Search'}
            'button_Queue' {$tip = 'Add to a corresponding queue.'}
            'button_Clear' {$tip = 'Clear text'}
            'button_Devices' {$tip = 'Moves selected search results to device queue'}
            'button_Groups' {$tip = 'Adds queued devices to all queued groups'}
            'OUBox' {$tip = 'Clear OU Queue'}
        }
        $search_tip.SetToolTip($this,$tip)
    }#end

#FORM
    #Creates Form
    $Form_ADGroups = New-Object System.Windows.Forms.Form
        $Form_ADGroups.Text = "PC Deployment Tool"
        $Form_ADGroups.Size = New-Object System.Drawing.Size(750,500)
        $Form_ADGroups.AutoSize = $true
        $Form_ADGroups.FormBorderStyle = "FixedDialog"
        $Form_ADGroups.TopMost = $true
        $Form_ADGroups.MaximizeBox = $false
        $Form_ADGroups.MinimizeBox = $true
        $Form_ADGroups.ControlBox = $true
        $Form_ADGroups.StartPosition = "CenterScreen"
        $Form_ADGroups.Font = "Arial"
        $Form_ADGroups.Icon = $form_ico
        
        
    
    #GroupBox for radio buttons
    $rad_grpbox = New-Object System.Windows.Forms.GroupBox
        $rad_grpbox.Top = '30'
        $rad_grpbox.Left = '95' #  ~= Form width - groupbox width / 2
        $rad_grpbox.Size = '540,75'
        $rad_grpbox.Text = 'Select a search option'

    #Label for searchbar
    $label_ADGroups = New-Object System.Windows.Forms.Label
        $label_ADGroups.Location = New-Object System.Drawing.Size(30,45)
        $label_ADGroups.AutoSize = $true
        $label_ADGroups.TextAlign = "MiddleCenter"
        $label_ADGroups.Text = "Name:"
        $Form_ADGroups.Controls.Add($label_ADGroups)

    #Searchbar
    $ADText = New-Object System.Windows.Forms.TextBox
        $ADText.Left = '70'
        $ADText.Top = '45' 
        $ADText.Size = '350,50'
        $ADText_Keydown = [System.Windows.Forms.KeyEventHandler]{
            if($_.KeyCode -eq 'Enter') 
            {
                $button_Search.PerformClick()
                $_.SuppressKeypress = $true
            }
        }#Handles Enter key the same way as clicking search
        $ADText.add_KeyDown($ADText_Keydown)
        $Form_ADGroups.Controls.Add($ADText)
    
    #DataQueue
        $DataBox = New-Object System.Windows.Forms.DataGridView
        $DataBox.Top = '180'
        $DataBox.Left = '20'
        $DataBox.Size = '180, 180'
        $DataBox.AutoSizeColumnsMode = 6
        $DataBox.ColumnCount = 1
        $DataBox.AllowUserToAddRows = 0
        $DataBox.Columns[0].MinimumWidth = $DataBox.Size.Width
        $DataBox.ColumnHeadersVisible = $true
        $DataBox.RowHeadersVisible = $false
        #$DataBox.SelectionMode = 0
        $DataBox.Columns[0].HeaderText = 'Name'
        $Form_ADGroups.Controls.Add($DataBox)

    #Device Queue
    $DeviceBox = New-Object System.Windows.Forms.DataGridView
        $DeviceBox.Top = 180
        $DeviceBox.Left = 280
        $DeviceBox.size = '180, 100'
        $DeviceBox.AutoSizeColumnsMode = 6
        $DeviceBox.ColumnCount = 1
        $DeviceBox.AllowUserToAddRows = 0
        $DeviceBox.Columns[0].MinimumWidth = $DeviceBox.Size.Width
        $DeviceBox.ColumnHeadersVisible = $true
        $DeviceBox.RowHeadersVisible = $false
        $DeviceBox.Columns[0].HeaderText = 'Device Name'
        $Form_ADGroups.Controls.Add($DeviceBox)

    #OU Queue    
    $OUBox = New-Object System.Windows.Forms.DataGridView
        $OUBox.Top = 300
        $OUBox.Left = 280
        $OUBox.size = '180, 60'
        $OUBox.AutoSizeColumnsMode = 6
        $OUBox.ColumnCount = 1
        $OUBox.AllowUserToAddRows = 0
        $OUBox.Columns[0].MinimumWidth = $OUBox.Size.Width
        $OUBox.ColumnHeadersVisible = $true
        $OUBox.RowHeadersVisible = $false
        $OUBox.Columns[0].HeaderText = 'OU'
        $Form_ADGroups.Controls.Add($OUBox)

###Search Radio buttons    
    #Group Queue
    $GroupBox = New-Object System.Windows.Forms.DataGridView
        $GroupBox.Top = $DataBox.Top
        $GroupBox.Left = ($DeviceBox.Left) + 260
        $GroupBox.size = $DataBox.Size
        $GroupBox.AutoSizeColumnsMode = 6
        $GroupBox.ColumnCount = 1
        $GroupBox.AllowUserToAddRows = 0
        $GroupBox.Columns[0].MinimumWidth = $DeviceBox.Size.Width
        $GroupBox.ColumnHeadersVisible = $true
        $GroupBox.RowHeadersVisible = $false
        $GroupBox.Columns[0].HeaderText = 'Group(s)'
        $Form_ADGroups.Controls.Add($GroupBox)


    #Radio for OU
    $OURadio = New-Object System.Windows.Forms.RadioButton
        $OURadio.Location = '30,16'
        $OURadio.AutoSize = $true
        $OURadio.Text = 'OU'
        $OURadio.Name = 'OURadio'
        $OURadio.add_MouseHover($ShowTip)
        $OURadio.Add_Click({radioevent})
        $Form_ADGroups.Controls.Add($OURadio)

    #Radio for Group Attributes
    $GARadio = New-Object System.Windows.Forms.RadioButton
        $GARadio.Location = '100,16'
        $GARadio.AutoSize = $true
        $GARadio.Text = 'Group Attributes'
        $GARadio.Name = 'GARadio'
        $GARadio.add_MouseHover($ShowTip)
        $GARadio.Add_Click({radioevent})
        $Form_ADGroups.Controls.Add($GARadio)

    #Radio for Device Name
    $DNRadio = New-Object System.Windows.Forms.RadioButton
        $DNRadio.Location = '220,16'
        $DNRadio.AutoSize = $true
        $DNRadio.Text = 'Device Name'
        $DNRadio.Name = 'DNRadio'
        $DNRadio.add_MouseHover($ShowTip)
        $DNRadio.Add_Click({radioevent})
        $Form_ADGroups.Controls.Add($DNRadio)

###Clear radio buttons
    $OURadio2 = New-Object System.Windows.Forms.RadioButton
        $OURadio2.Left = 75
        $OURadio2.Top = 430
        $OURadio2.AutoSize = $true
        $OURadio2.Text = 'OU'
        $OURadio2.Name = 'OUBox'
        $OURadio2.add_MouseHover($ShowTip)
        $OURadio2.Add_Click({$button_Clear2.Name = 'OUBox'})
        $Form_ADGroups.Controls.Add($OURadio2)

    #Clear Radio for Group Attributes
    $GARadio2 = New-Object System.Windows.Forms.RadioButton
        $GARadio2.Left = $OURadio2.Right + 10
        $GARadio2.Top = $OURadio2.Top
        $GARadio2.AutoSize = $true
        $GARadio2.Text = 'Groups'
        $GARadio2.Name = 'GroupBox'
        $GARadio2.add_MouseHover($ShowTip)
        $GARadio2.Add_Click({$button_Clear2.Name = 'GroupBox'})
        $Form_ADGroups.Controls.Add($GARadio2)

    #Clear Radio for Device Name
    $DNRadio2 = New-Object System.Windows.Forms.RadioButton
        $DNRadio2.Left = $GARadio2.Right + 10
        $DNRadio2.Top = $GARadio2.Top
        $DNRadio2.AutoSize = $true
        $DNRadio2.Text = 'Devices'
        $DNRadio2.Name = 'DeviceBox'
        $DNRadio2.add_MouseHover($ShowTip)
        $DNRadio2.Add_Click({$button_clear2.name = 'DeviceBox'})
        $Form_ADGroups.Controls.Add($DNRadio2)


    #Label for buttons
    $label_buttons = New-Object System.Windows.Forms.Label
        $label_buttons.Top = $DataBox.Top
        $label_buttons.Left = $DataBox.Left - 10
        #$label_buttons.AutoSize = $true
        #$label_buttons.multiline = $true 
        $label_buttons.Size = New-Object System.Drawing.Size(100,20)
        $label_buttons.TextAlign = "MiddleCenter"
        $label_buttons.Text = "Search results:"
        $Form_ADGroups.Controls.Add($label_buttons)

    #Search Button
    $button_Search = New-Object System.Windows.Forms.Button
        $button_Search.Top = '13'
        $button_Search.Left = '450'
        $button_Search.Size = '45,24'
        $button_Search.AutoSize = $true
        $button_Search.Name = 'button_Search'
        $button_Search.add_MouseHover($ShowTip)
        $button_Search.Image = $search_ico
        $button_Search.Add_Click({search_click})	
        $Form_ADGroups.Controls.Add($button_Search)

    #Clear text button
    $button_Clear = New-Object System.Windows.Forms.Button
        $button_Clear.Top = '45'
        $button_Clear.Left = '450'
        $button_Clear.Size = '45,24'
        $button_Clear.AutoSize = $true
        $button_Clear.Name = 'button_Clear'
        $button_Clear.add_MouseHover($ShowTip)
        $button_Clear.TextAlign = 'MiddleCenter'
        $button_Clear.Image = $Clear_ico
        $button_Clear.Add_Click({ClearView})
        $Form_ADGroups.Controls.Add($button_Clear)

    #Clear Devices
    $button_Clear2 = New-Object System.Windows.Forms.Button
        $button_Clear2.Top = $Form_ADGroups.Height - 75 
        $button_Clear2.Left = $DataBox.Left
        $button_Clear2.Size = '30,24'
        $button_Clear2.AutoSize = $true
        $button_Clear2.Name = ''
        $button_Clear2.add_MouseHover($ShowTip)
        $button_Clear2.TextAlign = 'MiddleCenter'
        $button_Clear2.Text = 'Clear'
        $button_Clear2.Add_Click({ClearView})
        $Form_ADGroups.Controls.Add($button_Clear2)

    #Queue button
    $button_Queue = New-Object System.Windows.Forms.Button
        $button_Queue.Top = 130
        $button_Queue.Left = $DataBox.Left + 60
        $button_Queue.Size = '45,24'
        $button_Queue.AutoSize = $true
        $button_Queue.Name = 'button_Queue'
        $button_Queue.add_MouseHover($ShowTip)
        $button_Queue.TextAlign = 'MiddleCenter'
        $button_Queue.Image = $queue_ico
        $button_Queue.Add_Click({move_selected})
        $Form_ADGroups.Controls.Add($button_Queue)
        


    #Adds radios to groupbox and groupbox to the form
        $rad_grpbox.Controls.AddRange(@($OURadio, $GARadio, $DNRadio, $label_ADGroups, $ADText, $button_Search, $button_Clear))
        $Form_ADGroups.Controls.Add($rad_grpbox)



    #OU Button; Moves search results to device queue
    $button_Devices = New-Object System.Windows.Forms.Button
        #$button_Devices.FlatStyle = 'Standard'
        $button_Devices.Top = 130
        $button_Devices.Left = $DeviceBox.Left + 40
        #$button_Devices.Anchor = 'right,bottom'
        $button_Devices.Size = New-Object System.Drawing.Size(100,24)
        $button_Devices.TextAlign = "MiddleCenter"
        $button_Devices.add_MouseHover($ShowTip)
        $button_Devices.Text = "Move to OU"
        #On-Click
            $button_Devices.Add_Click({moveOU_click})
        $Form_ADGroups.Controls.Add($button_Devices)


    #Group Button; Adds results group queue
    $button_Groups = New-Object System.Windows.Forms.Button
        $button_Groups.Top = 130
        $button_Groups.Left = $GroupBox.Left + 40
        $button_Groups.FlatStyle = 'Standard'
        #$button_Groups.Anchor = 'right,bottom'
        $button_Groups.Size = New-Object System.Drawing.Size(100,24)
        $button_Groups.TextAlign = "MiddleCenter"
        $button_Groups.add_MouseHover($ShowTip)
        $button_Groups.Text = "Add Groups"
        #On-Click
            $button_Groups.Add_Click({addgrp_click})
        $Form_ADGroups.Controls.Add($button_Groups)

#Show Form
    $Form_ADGroups.Topmost = $true
    $Form_ADGroups.Add_Shown({$Form_ADGroups.Activate()})
    [void]$Form_ADGroups.ShowDialog()


