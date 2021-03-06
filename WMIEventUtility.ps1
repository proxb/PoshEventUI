#Load Required Assemblies
Add-Type –assemblyName PresentationFramework
Add-Type –assemblyName PresentationCore
Add-Type –assemblyName WindowsBase
Add-Type –assemblyName Microsoft.VisualBasic
Add-Type –assemblyName System.Windows.Forms

$Script:Computername = $Env:Computername

#Ensure that we are running the GUI from the correct location
Set-Location $(Split-Path $MyInvocation.MyCommand.Path)

##Functions
Function Clear-ListView {
    $ConsumerListView.Items.Clear()
    $FilterListView.Items.Clear()
    $BindingListView.Items.Clear()
}

Function Add-Computername {
    $Script:Computername = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a computer to connect to", "Connect to Computer",$Env:Computername)
    If (-Not [string]::IsNullOrEmpty($Computername)) {
        If (Test-Connection -ComputerName $Computername -Count 1 -Quiet) {
            $StatusTextbox.Text = ("Attempting connection to {0}..." -f $Computername)
            Clear-ListView
            Get-EventSubscription -Computername $Computername
            $StatusTextbox.Text = "Connected to $Computername"
        }
    }
    $ConsumerComboBox.SelectedIndex = 0
}

Function Get-EventSubscription {
    [cmdletbinding()]
    Param ($Computername = $Env:Computername)
    #Event Consumers
    $Script:EventConsumers = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __EventConsumer | Sort Name
    $EventConsumers | ForEach {
        $ConsumerListView.Items.Add($_.Name)
    }

    #Event Filters
    $Script:EventFilters = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __EventFilter | Sort Name
    $EventFilters | ForEach {
        $FilterListView.Items.Add($_.Name)
    }

    #Event Bindings
    $Script:EventBindings = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __FilterToConsumerBinding | Sort __Path
    $EventBindings | ForEach {
        $BindingListView.Items.Add($_.__Path)
    }
}

#Build the GUI
[xml]$xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='Window' Title='WMI Event Utility' WindowStartupLocation = 'CenterScreen' 
    Width = '865' Height = '575' ShowInTaskbar = 'True'>
        <Window.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Window.Background> 
    <Grid x:Name = 'Grid' ShowGridLines = 'false'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = '*'/>
            <RowDefinition Height = '5'/>
            <RowDefinition Height = '25'/>
        </Grid.RowDefinitions>    
        <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
        <Menu.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Menu.Background>
            <MenuItem x:Name = 'FileMenu' Header = '_File'>
                <MenuItem x:Name = 'ConnectMenu' Header = 'C_onnect To Another Computer' ToolTip = 'Exits the utility.' InputGestureText ='Ctrl+C'> </MenuItem>
                <Separator />
                <MenuItem x:Name = 'ExitMenu' Header = 'E_xit' ToolTip = 'Exits the utility.' InputGestureText ='Ctrl+E'> </MenuItem>
            </MenuItem>           
        </Menu>                               
        <Grid Grid.Row = '1' ShowGridLines = 'false'>  
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>                
                <ColumnDefinition Width="10"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = '*'/>
                <RowDefinition Height = 'Auto'/>
            </Grid.RowDefinitions>  
            <Label Content='Filter' Grid.Column = '1' Grid.Row = '1' HorizontalAlignment = 'Center' 
                FontWeight = 'Bold' FontSize = '14'/>         
            <Label Content='Consumer' Grid.Column = '3' Grid.Row = '1' HorizontalAlignment = 'Center' 
                FontWeight = 'Bold' FontSize = '14'/>         
            <Label Content='Binding' Grid.Column = '5' Grid.Row = '1' HorizontalAlignment = 'Center' 
                FontWeight = 'Bold' FontSize = '14'/>   
            <ListView x:Name = 'FilterListView' Grid.Column = '1' Grid.Row = '2' ToolTip = 'Double click selected item to view more information'/>
            <Grid Grid.Column = '3' Grid.Row = '2' ShowGridLines = 'false'>             
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = 'Auto'/>
                <RowDefinition Height = '5'/>
                <RowDefinition Height = '*'/>
            </Grid.RowDefinitions>   
            <ComboBox x:Name = 'ConsumerComboBox' Grid.Row = '0' MaxWidth = '200' IsReadOnly = 'True' SelectedIndex = '0'>
                <TextBlock> All </TextBlock>
                <TextBlock> ActiveScriptEventConsumer </TextBlock>
                <TextBlock> CommandLineEventConsumer </TextBlock> 
                <TextBlock> LogFileEventConsumer </TextBlock>
                <TextBlock> NTEventLogEventConsumer </TextBlock> 
                <TextBlock> SMTPEventConsumer </TextBlock>    
            </ComboBox>                                     
            <ListView x:Name = 'ConsumerListView' Grid.Row = '2' ToolTip = 'Double click selected item to view more information'/>      
            </Grid>                 
            <ListView x:Name = 'BindingListView' Grid.Column = '5' Grid.Row = '2' ToolTip = 'Double click selected item to view more information'/>  
            <Grid Grid.Column = '1' Grid.Row = '3' ShowGridLines = 'False'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height = '5'/>
                    <RowDefinition Height = '*'/>
                </Grid.RowDefinitions>  
                <Button x:Name = 'NewFilterButton' Grid.Column = '0' Grid.Row = '1' Content = 'New' MaxWidth = '125'/>           
                <Button x:Name = 'RemoveFilterButton' Grid.Column = '2' Grid.Row = '1' Content = 'Remove' MaxWidth = '125'/>                            
            </Grid>   
            <Grid Grid.Column = '3' Grid.Row = '3' ShowGridLines = 'False'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height = '5'/>
                    <RowDefinition Height = '*'/>
                </Grid.RowDefinitions>                   
                <Button x:Name = 'NewConsumerButton' Grid.Column = '0' Grid.Row = '1' Content = 'New' MaxWidth = '125'/>           
                <Button x:Name = 'RemoveConsumerButton' Grid.Column = '2' Grid.Row = '1' Content = 'Remove' MaxWidth = '125'/>                            
            </Grid> 
            <Grid Grid.Column = '5' Grid.Row = '3' ShowGridLines = 'False'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height = '5'/>
                    <RowDefinition Height = '*'/>
                </Grid.RowDefinitions>  
                <Button x:Name = 'NewBindingButton' Grid.Column = '0' Grid.Row = '1' Content = 'New' MaxWidth = '125'/>           
                <Button x:Name = 'RemoveBindingButton' Grid.Column = '2' Grid.Row = '1' Content = 'Remove' MaxWidth = '125'/>           
            </Grid>                         
        </Grid>  
        <Separator Grid.Row = '2' />  
        <TextBox x:Name = 'StatusTextbox' Grid.Row = '3' IsReadOnly = 'True' Text = 'Connected to $Computername'>       
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

#regionConnect To Controls
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Global:Window=[Windows.Markup.XamlReader]::Load( $reader )


$ConsumerListView = $Window.FindName('ConsumerListView')
$FilterListView = $Window.FindName('FilterListView')
$BindingListView = $Window.FindName('BindingListView')
$RemoveFilterButton = $Window.FindName('RemoveFilterButton')
$RemoveConsumerButton = $Window.FindName('RemoveConsumerButton')
$RemoveBindingButton = $Window.FindName('RemoveBindingButton')
$NewFilterButton = $Window.FindName('NewFilterButton')
$NewConsumerButton = $Window.FindName('NewConsumerButton')
$NewBindingButton = $Window.FindName('NewBindingButton')
$ExitMenu = $Window.FindName('ExitMenu')
$ConnectMenu = $Window.FindName('ConnectMenu')
$StatusTextbox = $Window.FindName('StatusTextbox')
$ConsumerComboBox = $Window.FindName('ConsumerComboBox')
#endregion

#region Events
#Initial Startup
$Window.Add_Loaded({
    #Clear all list views
    Clear-ListView
    
    #Get event consumers from local system
    Get-EventSubscription 
})

#region Remove WMI Events
#RemoveFilterButton
$RemoveFilterButton.Add_Click({
    $Filters = $FilterListView.SelectedItems | ForEach {$_}
    ForEach ($item in $Filters) {
        $EventFilters | Where {
            $_.Name -eq $item
        } | Remove-WmiObject
        $FilterListView.Items.Remove($item)
    }
})

#RemoveConsumerButton
$RemoveConsumerButton.Add_Click({
    $Consumers = $ConsumerListView.SelectedItems | ForEach {$_}
    ForEach ($item in $Consumers) {
        $EventConsumers | Where {
            $_.Name -eq $item
        } | Remove-WmiObject
        $ConsumerListView.Items.Remove($item)
    }
})

#RemoveBindingButton
$RemoveBindingButton.Add_Click({
    $Bindings = $BindingListView.SelectedItems | ForEach {$_}
    ForEach ($item in $Bindings) {
        $EventBindings | Where {
            $_.__Path -eq $item
        } | Remove-WmiObject
        $BindingListView.Items.Remove($item)
    }
})
#endregion

#region Create WMI Events
#Create Filter Button
$NewFilterButton.Add_Click({
    $Return = .\Create-WMIFilter.ps1 -Computername $Computername
    If ($Return) {
        $FilterListView.Items.Clear()
        $Script:EventFilters = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __EventFilter | Sort Name
        $EventFilters | ForEach {
            $FilterListView.Items.Add($_.Name)
        }     
    }
})

#Create Binding Button
$NewBindingButton.Add_Click({
    $Return = .\Create-WMIBinding.ps1 -Computername $Computername -Filter $Script:EventFilters -Consumer $Script:EventConsumers
    If ($Return) {
        $BindingListView.Items.Clear()
        $Script:EventBindings = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __FilterToConsumerBinding | Sort __Path
        $EventBindings | ForEach {
            $BindingListView.Items.Add($_.__Path)
        }        
    }
})

#Create Consumer Button
$NewConsumerButton.Add_Click({
    $Return = .\Create-WMIConsumer.ps1 -Computername $Computername
    If ($Return) {
        $ConsumerListView.Items.Clear()
        $Script:EventConsumers = Get-WMIObject -Computername $Computername -Namespace root\Subscription -Class __EventConsumer | Sort __Path
        $EventConsumers | ForEach {
            $ConsumerListView.Items.Add($_.Name)
        }        
    }
})
#endregion

#Exit Menu
$ExitMenu.Add_Click({
    $Global:Window.Close()
})

#ConnectToComputer Menu
$ConnectMenu.Add_Click({
    Add-Computername
})

#FilterListView Doubleclick
$FilterListView.Add_MouseDoubleClick({
    If (-Not [string]::IsNullOrEmpty($this.SelectedItem)) {
        $EventFilters | Where {
            $_.Name -eq $this.SelectedItem
        } | Select __Server,__Path,Name,Query,QueryLanguage | Out-GridView -Title EventFilters
    }
})

#ConsumerListView Doubleclick
$ConsumerListView.Add_MouseDoubleClick({
$item = $EventConsumers | Where {
    $_.Name -eq $This.SelectedItem
}
Switch ($item.__Class) {
    'CommandLineEventConsumer' {
        $item | Select __Server,__Path,Name,__Class,CommandLineTemplate,ExecutablePath,WorkingDirectory,DesktopName,MachineName | 
            Out-GridView -Title EventConsumers
    }    
    'LogFileEventConsumer' {
        $item | Select __Server,__Path,Name,__Class,FileName,Text,MaximumFileSize | 
            Out-GridView -Title EventConsumers    
    }
    'ActiveScriptEventConsumer' {
        $item | Select __Server,__Path,Name,__Class,ScriptFileName,ScriptingEngine,ScriptText,KillTimeout | 
            Out-GridView -Title EventConsumers     
    }
    'SMTPEventConsumer' {
        $item | Select __Server,__Path,Name,__Class,ToLine,CcLine,Bccline,Subject,FromLine,Message,HeaderFields,ReplyToLine,SMTPAddress | 
            Out-GridView -Title EventConsumers     
    }
    'NTEventLogEventConsumer' {
        $item | Select __Server,__Path,Name,__Class,SourceName,InsertionStringTemplates,NameOrRawDataProperty,NumberOfInsertionStrings,EventType,Category | 
            Out-GridView -Title EventConsumers
    }
}
})

#BindingListView Doubleclick
$BindingListView.Add_MouseDoubleClick({
    If (-Not [string]::IsNullOrEmpty($this.SelectedItem)) {
        $EventBindings | Where {
            $_.__Path -eq $this.SelectedItem
        } | Select __Server,__Path,Filter,Consumer | Out-GridView -Title EventBindings
    }
})

#Combobox Change Events
$ConsumerComboBox.Add_SelectionChanged({
    $ConsumerListView.Items.Clear()
    If ($This.SelectedItem.Text -eq 'All') {
        $EventConsumers | ForEach {
            $ConsumerListView.Items.Add($_.Name)
        }
    } Else {
        $EventConsumers | Where {
            $_.__Class -eq $This.SelectedItem.Text
        } | ForEach {
            $ConsumerListView.Items.Add($_.Name)
        }
    }
})

#Key events
$Window.Add_KeyDown({ 
    $key = $_.Key  
    If ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) {
        Switch ($Key) {
        "E" {$This.Close()}
        "C" {Add-Computername}
        Default {$Null}
        }
    }   
})
#endregion
##Display UI
[void]$Global:Window.ShowDialog()