Param ($Computername = $Env:Computername)
#region Exclusion Collection
$Script:Exclusion = New-Object System.Collections.ArrayList
$Exclusion.AddRange(("consumerLabel","ConsumerComboBox","nameLabel","eventNamespaceLabel","Name_txtbx","Namespace_txtbx","childGrid","StatusTextbox")) | Out-Null
$Script:exclusionCount = $Exclusion.count
#endregion

#region Functions
    #region Consumer UI Setup
Function Set-ActiveScriptWindow {
    $Script:Exclusion.AddRange(("label1_1","label1_2","label4_1","textbox1_1","textbox4_1","scriptlanguagecombobox")) | Out-Null
    $tempExclusionCount = $Exclusion.count
    $label1_1.Content = "Script File"
    $label1_2.Content = "Scripting Engine"
    $label4_1.Content = "Script Text"
    $grid.Children | ForEach {
        If ($Exclusion -notcontains $_.name) {
            $_.Visibility = "Hidden"
        } Else {
            $_.Visibility = "Visible"
        }
    }
    $Script:Exclusion.RemoveRange($exclusionCount,($tempExclusionCount - $ExclusionCount))
}

Function Set-CommandLineWindow {
    $Script:Exclusion.AddRange(("label1_1","label1_2","label4_1","textbox1_1","textbox1_2","textbox4_1")) | Out-Null
    $tempExclusionCount = $Exclusion.count
    $label1_1.Content = "Working Directory"
    $label1_2.Content = "Executable Path"
    $label4_1.Content = "Command Line Template"
    $grid.Children | ForEach {
        If ($Exclusion -notcontains $_.name) {
            $_.Visibility = "Hidden"
        } Else {
            $_.Visibility = "Visible"
        }
    }
    $Script:Exclusion.RemoveRange($exclusionCount,($tempExclusionCount - $ExclusionCount))
}

Function Set-LogFileWindow {
    $Script:Exclusion.AddRange(("label1_1","label1_2","label1_3","label4_1","textbox1_1","textbox1_2","interactiveCheckBox","textbox4_1")) | Out-Null
    $tempExclusionCount = $Exclusion.count
    $label1_1.Content = "File Name"
    $label1_2.Content = "Maximum File Size (0 - 65535)"
    $label1_3.Content = "Encoding"
    $interactiveCheckBox.Content = "IsUnicode"
    $label4_1.Content = "File Text"
    $grid.Children | ForEach {
        If ($Exclusion -notcontains $_.name) {
            $_.Visibility = "Hidden"
        } Else {
            $_.Visibility = "Visible"
        }
    }
    $Script:Exclusion.RemoveRange($exclusionCount,($tempExclusionCount - $ExclusionCount))
}

Function Set-NTEventLogWindow {
    $Script:eventType = @{
        Success =      '0x0000' #EVENTLOG_SUCCESS
        Error =        '0x0001' #EVENTLOG_ERROR
        Warning =      '0x0002' #EVENTLOG_WARNING
        Information =  '0x0004' #EVENTLOG_INFORMATION_TYPE
        AuditSuccess = '0x0008' #EVENTLOG_AUDIT_SUCCESS
        AuditFailure = '0x0010' #EVENTLOG_AUDIT_FAILURE
    }
    $Script:Exclusion.AddRange(("label1_1","label1_2","label1_3","label2_1","label2_2","label4_1","textbox1_1","textbox1_2","textbox1_3",
    "textbox2_1","textbox4_1","eventtypecombobox")) | Out-Null
    $tempExclusionCount = $Exclusion.count
    $label1_1.Content = "Source"
    $label1_2.Content = "Event ID"
    $label1_3.Content = "Category"
    $label2_1.Content = "UNC Server Name"
    $label2_2.Content = "Event Type"
    $label4_1.Content = "Insertion String Templates"
    $grid.Children | ForEach {
        If ($Exclusion -notcontains $_.name) {
            $_.Visibility = "Hidden"
        } Else {
            $_.Visibility = "Visible"
        }
    }
    $Script:Exclusion.RemoveRange($exclusionCount,($tempExclusionCount - $ExclusionCount))
}

Function Set-SMTPEventLogWindow {
    $Script:Exclusion.AddRange(("label1_1","label1_2","label1_3","label2_1","label2_2","label2_3","label3_1","label3_2","label4_1","textbox1_1",
    "textbox1_2","textbox1_3","textbox2_1","textbox2_2","textbox2_3","textbox3_1","textbox3_2","textbox4_1")) | Out-Null
    $tempExclusionCount = $Exclusion.count
    $label1_1.Content = "To"
    $label1_2.Content = "From"
    $label1_3.Content = "Subject"
    $label2_1.Content = "Cc"
    $label2_2.Content = "Bcc"
    $label2_3.Content = "Reply To Line"
    $label3_1.Content = "SMTPServer"
    $label3_2.Content = "HeaderFields (separate by ',' or ';')"
    $label4_1.Content = "Message"
    $grid.Children | ForEach {
        If ($Exclusion -notcontains $_.name) {
            $_.Visibility = "Hidden"
        } Else {
            $_.Visibility = "Visible"
        }
    }
    $Script:Exclusion.RemoveRange($exclusionCount,($tempExclusionCount - $ExclusionCount))
}
    #endregion
    #region Consumer Creation
    Function New-ActiveScriptConsumer {
        Param ($Computername)
        $instanceConsumer = ([wmiclass]"\\$Computername\root\subscription:ActiveScriptEventConsumer").CreateInstance()
        $instanceConsumer.Name = $Name_txtbx.text
        $instanceConsumer.ScriptingEngine = $scriptlanguagecombobox.Text
        $instanceConsumer.ScriptFilename = $textbox1_1.text.Trim()
        $instanceConsumer.ScriptText = $textbox4_1.text.Trim()
        Try {
            If ($instanceConsumer.ScriptFilename -AND $instanceConsumer.ScriptText) {
                Throw "You can only specify either the ScriptFileName or ScriptText!"
            } ElseIf (-Not ($instanceConsumer.ScriptFilename -OR $instanceConsumer.ScriptText)) {
                Throw "You must specify either ScriptFileName or ScriptText!"
            }
            #Create the WMI Event Consumer
            $instanceConsumer.Put()                
            $Consumer_Window.DialogResult = $True
            $Consumer_Window.Close()
        } Catch {
            $StatusTextbox.Foreground = 'Red'
            $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
        }
    }
    Function New-CommandLineConsumer {
        Param ($Computername)
        $instanceConsumer = ([wmiclass]"\\$Computername\root\subscription:CommandLineEventConsumer").CreateInstance()
        $instanceConsumer.Name = $Name_txtbx.text
        $instanceConsumer.CommandLineTemplate = $textbox4_1.text.Trim()
        $instanceConsumer.WorkingDirectory = $textbox1_1.text.Trim()
        $instanceConsumer.ExecutablePath = $textbox1_2.text.Trim()
        Try {
            #Create the WMI Event Consumer
            $instanceConsumer.Put()                
            $Consumer_Window.DialogResult = $True
            $Consumer_Window.Close()
        } Catch {
            $StatusTextbox.Foreground = 'Red'
            $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
        }
    }
    Function New-LogFileConsumer {
        Param ($Computername)
        $instanceConsumer = ([wmiclass]"\\$Computername\root\subscription:LogFileEventConsumer").CreateInstance()
        $instanceConsumer.Filename = $textbox1_1.text
        $instanceConsumer.Name = $Name_txtbx.text
        $instanceConsumer.Text = $textbox4_1.text
        If ($textbox1_2.text.length -gt 0) {
            $instanceConsumer.MaximumFileSize = $textbox1_2.text
        }
        If ($interactiveCheckBox.IsChecked) {
            $instanceConsumer.IsUnicode = $True
        }
        Try {
            #Create the WMI Event Consumer
            $instanceConsumer.Put()                
            $Consumer_Window.DialogResult = $True
            $Consumer_Window.Close()
        } Catch {
            $StatusTextbox.Foreground = 'Red'
            $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
        }
    }
    Function New-NTEventLogConsumer {
        Param ($Computername)
        $instanceConsumer = ([wmiclass]"\\$Computername\root\subscription:NTEventLogEventConsumer").CreateInstance()
        $instanceConsumer.Name = $Name_txtbx.text
        If ($textbox1_1.text.length -gt 0) {
            $instanceConsumer.SourceName = $textbox1_1.text
        }
        If ($textbox1_2.text.length -gt 0) {
            $instanceConsumer.EventID = $textbox1_2.text
        }
        If ($textbox1_3.text.length -gt 0) {
            $instanceConsumer.Category = $textbox1_3.text
        } 
        If ($textbox2_1.text.length -gt 0) {
            $instanceConsumer.UNCServerName = $textbox2_1.text
        }
        If ($textbox4_1.text.length -gt 0) {
            $instanceConsumer.InsertionStringTemplates = ($textbox4_1.text -split "\n")
        } Else {
            Throw "You must specify an Insertion String!"
        }
        $instanceConsumer.NumberOfInsertionStrings = $instanceConsumer.InsertionStringTemplates.Count
        $instanceConsumer.EventType = $eventType[$eventtypecombobox.Text]
        Try {
            write-verbose "$($instanceConsumer | out-string)" -Verbose
            Write-Verbose ("Number of insertion strings: {0}" -f ($instanceConsumer.InsertionStringTemplates).Count) -Verbose
            #Create the WMI Event Consumer
            $instanceConsumer.Put()                
            $Consumer_Window.DialogResult = $True
            $Consumer_Window.Close()
        } Catch {
            $StatusTextbox.Foreground = 'Red'
            $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
        }
    }
    Function New-SMTPEventLogConsumer {
        Param ($Computername)
        $instanceConsumer = ([wmiclass]"\\$Computername\root\subscription:SMTPEventConsumer").CreateInstance()
        $instanceConsumer.Name = $Name_txtbx.text.Trim()
        If ($textbox2_2.text.Trim().length -gt 0) {
            $instanceConsumer.BccLine = $textbox2_2.text.Trim()
        }
        If ($textbox2_1.text.Trim().length -gt 0) {
            $instanceConsumer.CcLine = $textbox2_1.text.Trim()
        }
        If ($textbox1_2.text.Trim().length -gt 0) {
            $instanceConsumer.FromLine = $textbox1_2.text.Trim()
        }
        $instanceConsumer.Message = $textbox4_1.text.Trim()
        $instanceConsumer.ReplyToLine = $textbox2_3.text.Trim()
        $instanceConsumer.Subject = $textbox1_3.text.Trim()
        If ($textbox1_1.text.Trim().length -gt 0) {
            $instanceConsumer.ToLine = $textbox1_1.text.Trim()
        }
        If ($textbox3_2.text.Trim().length -gt 0) {
            $instanceConsumer.HeaderFields = ($textbox3_2.text.Trim() -split ",|;")
        }
        $instanceConsumer.SMTPServer = $textbox3_1.text.Trim()
        Try {
            #Create the WMI Event Consumer
            $instanceConsumer.Put()                
            $Consumer_Window.DialogResult = $True
            $Consumer_Window.Close()
        } Catch {
            $StatusTextbox.Foreground = 'Red'
            $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
        }
    }
    #endregion
#endregion

#region Build the GUI
[xml]$consumer_xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='Window' Title='New WMI Event Consumer on $Computername' WindowStartupLocation = 'CenterScreen' 
    SizeToContent = 'Height' Width = '750' ShowInTaskbar = 'True' ResizeMode = 'Noresize'>
        <Window.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Window.Background> 
    <Grid x:Name = 'Grid' ShowGridLines = 'false'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = '*'/>            
            <RowDefinition Height = 'Auto'/>  
        </Grid.RowDefinitions>  
        <Label x:Name = 'consumerLabel' Content='Select Consumer Type' Grid.Column = '0' Grid.Row = '0' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>
        <ComboBox x:Name = 'ConsumerComboBox' Grid.Row = '1' Grid.Column = '0' MaxWidth = '200' IsReadOnly = 'True' 
        HorizontalAlignment = 'Left'>
            <TextBlock> ActiveScriptEventConsumer </TextBlock>
            <TextBlock> CommandLineEventConsumer </TextBlock> 
            <TextBlock> LogFileEventConsumer </TextBlock>
            <TextBlock> NTEventLogEventConsumer </TextBlock> 
            <TextBlock> SMTPEventConsumer </TextBlock>            
        </ComboBox> 
        <Label x:Name = 'nameLabel' Content='Name' Grid.Column = '2' Grid.Row = '0' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>
        <Label x:Name = 'eventNamespaceLabel' Content='EventNamespace' Grid.Column = '4' Grid.Row = '0' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>
        <Label x:Name = 'label1_1' Content='NULL1-1' Grid.Column = '0' Grid.Row = '4' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>  
        <Label x:Name = 'label1_2' Content='NULL1-2' Grid.Column = '2' Grid.Row = '4' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/> 
        <Label x:Name = 'label1_3' Content='NULL1-3' Grid.Column = '4' Grid.Row = '4' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>  
        <Label x:Name = 'label2_1' Content='NULL2-1' Grid.Column = '0' Grid.Row = '6' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>  
        <Label x:Name = 'label2_2' Content='NULL2-2' Grid.Column = '2' Grid.Row = '6' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/> 
        <Label x:Name = 'label2_3' Content='NULL2-3' Grid.Column = '4' Grid.Row = '6' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>  
        <Label x:Name = 'label3_1' Content='NULL3-1' Grid.Column = '0' Grid.Row = '8' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>   
        <Label x:Name = 'label3_2' Content='NULL3-2' Grid.Column = '2' Grid.Row = '8' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>                                                                                                          
        <Label x:Name = 'label4_1' Content='NULL4-1' Grid.Column = '0' Grid.Row = '10' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>   
        <TextBox x:Name = 'Name_txtbx' Grid.Row = '1' Grid.Column = '2' Text = 'ConsumerName'/>         
        <TextBox x:Name = 'Namespace_txtbx' Grid.Row = '1' Grid.Column = '6' Text = 'root/CIMV2' />
        <TextBox x:Name = 'textbox1_1' Grid.Row = '5' Grid.Column = '0' />
        <TextBox x:Name = 'textbox1_2' Grid.Row = '5' Grid.Column = '2' />
        <TextBox x:Name = 'textbox1_3' Grid.Row = '5' Grid.Column = '4' />   
        <TextBox x:Name = 'textbox2_1' Grid.Row = '7' Grid.Column = '0' />
        <TextBox x:Name = 'textbox2_2' Grid.Row = '7' Grid.Column = '2' />
        <TextBox x:Name = 'textbox2_3' Grid.Row = '7' Grid.Column = '4' />          
        <TextBox x:Name = 'textbox3_1' Grid.Row = '9' Grid.Column = '0' />   
        <TextBox x:Name = 'textbox3_2' Grid.Row = '9' Grid.Column = '2' /> 
        <TextBox x:Name = 'textbox4_1' Grid.Row = '11' Grid.Column = '0' Grid.ColumnSpan = '5' TextWrapping = 'Wrap' AcceptsReturn = 'True'
            MinLines = '6'/> 
        <CheckBox x:Name = 'interactiveCheckBox' Grid.Row = '5' Grid.Column = '4'/>        
        <ComboBox x:Name = 'scriptlanguagecombobox' Grid.Row = '5' Grid.Column = '2' MaxWidth = '100' IsReadOnly = 'True' SelectedIndex = '0'
        HorizontalAlignment = 'Left'>
            <TextBlock> VBScript </TextBlock>
        </ComboBox>   
        <ComboBox x:Name = 'eventtypecombobox' Grid.Row = '7' Grid.Column = '2' MaxWidth = '100' IsReadOnly = 'True' SelectedIndex = '0'
        HorizontalAlignment = 'Left'>
            <TextBlock> Success </TextBlock>
            <TextBlock> Failure </TextBlock>
            <TextBlock> Warning </TextBlock>
            <TextBlock> Information </TextBlock>
            <TextBlock> AuditSuccess </TextBlock>
            <TextBlock> AuditFailure </TextBlock>
        </ComboBox>                                                    
        <Grid x:Name='childGrid' Grid.Row = '12' Grid.Column = '2' ShowGridLines = 'False'>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = '5'/>       
                <RowDefinition Height = 'Auto'/>  
            </Grid.RowDefinitions>   
            <Button x:Name = 'CreateButton' Content ='Create' Grid.Column = '0' Grid.Row = '1' MaxWidth = '125' IsDefault = 'True' Margin = "0,0,0,5"/>       
            <Button x:Name = 'CancelButton' Content = 'Cancel' Grid.Column = '2' Grid.Row = '1' MaxWidth = '125' IsCancel = 'True' Margin = "0,0,0,5"/>       
        </Grid>
        <TextBox x:Name = 'StatusTextbox' Grid.Row = '13' Grid.ColumnSpan = '5' IsReadOnly = 'True'>       
            <TextBox.Background>
                <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                    <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                    <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
                </LinearGradientBrush>
            </TextBox.Background>         
        </TextBox>        
    </Grid>   
</Window>
"@ 
#endregion

#region Controls
$reader=(New-Object System.Xml.XmlNodeReader $consumer_xaml)
$Global:Consumer_Window=[Windows.Markup.XamlReader]::Load( $reader )
$StatusTextbox = $Consumer_Window.FindName('StatusTextbox')
$CreateButton = $Consumer_Window.FindName('CreateButton')
$CancelButton = $Consumer_Window.FindName('CancelButton')
$ConsumerComboBox = $Consumer_Window.FindName('ConsumerComboBox')
$grid = $Consumer_Window.FindName('Grid')
$Name_txtbx = $Consumer_Window.FindName('Name_txtbx')
$label1_1 = $Consumer_Window.FindName('label1_1')
$label1_2 = $Consumer_Window.FindName('label1_2')
$label1_3 = $Consumer_Window.FindName('label1_3')
$label2_1 = $Consumer_Window.FindName('label2_1')
$label2_2 = $Consumer_Window.FindName('label2_2')
$label2_3 = $Consumer_Window.FindName('label2_3')
$label3_1 = $Consumer_Window.FindName('label3_1')
$label3_2 = $Consumer_Window.FindName('label3_2')
$label4_1 = $Consumer_Window.FindName('label4_1')
$textbox1_1 = $Consumer_Window.FindName('textbox1_1')
$textbox1_2 = $Consumer_Window.FindName('textbox1_2')
$textbox1_3 = $Consumer_Window.FindName('textbox1_3')
$textbox2_1 = $Consumer_Window.FindName('textbox2_1')
$textbox2_2 = $Consumer_Window.FindName('textbox2_2')
$textbox2_3 = $Consumer_Window.FindName('textbox2_3')
$textbox3_1 = $Consumer_Window.FindName('textbox3_1')
$textbox3_2 = $Consumer_Window.FindName('textbox3_2')
$textbox4_1 = $Consumer_Window.FindName('textbox4_1')
$scriptlanguagecombobox = $Consumer_Window.FindName('scriptlanguagecombobox')
$interactiveCheckBox = $Consumer_Window.FindName('interactiveCheckBox')
$eventtypecombobox = $Consumer_Window.FindName('eventtypecombobox')
#endregion

#region Events
#Window Load
$Consumer_Window.Add_Activated({
    If (-Not $alreadyDone) {
        $ConsumerComboBox.SelectedIndex = '0'
        $This.UpdateLayout()
        $Script:alreadyDone = $True
    }
})

#Window Close
$Consumer_Window.Add_Closed({
    $Script:alreadyDone = $False
})
#Cancel
$CancelButton.Add_Click({    
    $Consumer_Window.Close()      
})

#Create
$CreateButton.Add_Click({
    Switch ($ConsumerComboBox.SelectedItem.Text) {
        "ActiveScriptEventConsumer" {
            New-ActiveScriptConsumer -Computername $Computername
        }
        "CommandLineEventConsumer" {
            New-CommandLineConsumer -Computername $Computername
        }
        "LogFileEventConsumer" {
            New-LogFileConsumer -Computername $Computername
        }
        "NTEventLogEventConsumer" {
            New-NTEventLogConsumer -Computername $Computername
        }
        "SMTPEventConsumer" {
            New-SMTPEventLogConsumer -Computername $Computername
        }
        
    }
})

#ComboBox Change
$ConsumerComboBox.Add_SelectionChanged({   
    Switch ($This.SelectedItem.Text) {
        "ActiveScriptEventConsumer" {
            Set-ActiveScriptWindow
        }
        "CommandLineEventConsumer" {
            Set-CommandLineWindow
        }
        "LogFileEventConsumer" {
            Set-LogFileWindow
        }
        "NTEventLogEventConsumer" {
            Set-NTEventLogWindow
        }
        "SMTPEventConsumer" {
            Set-SMTPEventLogWindow
        }
        
    }
})
#endregion


$Consumer_Window.ShowDialog()