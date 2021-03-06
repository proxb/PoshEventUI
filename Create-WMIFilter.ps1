Param ($Computername = $Env:Computername)

#Build the GUI
[xml]$filter_xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='Window' Title='New WMI Event Filter on $Computername' WindowStartupLocation = 'CenterScreen' 
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
            <RowDefinition Height = '*'/>            
            <RowDefinition Height = 'Auto'/>  
        </Grid.RowDefinitions>  
        <Label Content='Name' Grid.Column = '0' Grid.Row = '0' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>
        <Label Content='EventNamespace' Grid.Column = '2' Grid.Row = '0' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>
        <Label Content='QueryLanguage' Grid.Column = '4' Grid.Row = '0' HorizontalAlignment = 'Center' 
                FontWeight = 'Bold' FontSize = '12'/>  
        <Label Content='WQL Query' Grid.Column = '0' Grid.Row = '2' HorizontalAlignment = 'Left' 
                FontWeight = 'Bold' FontSize = '12'/>   
        <TextBox x:Name = 'Name_txtbx' Grid.Row = '1' Grid.Column = '0' Text = 'FilterName'/>         
        <TextBox x:Name = 'Namespace_txtbx' Grid.Row = '1' Grid.Column = '2' Text = 'root/CIMV2' />  
        <TextBox x:Name = 'Query_txtbx' Grid.Row = '3' Grid.Column = '0' Grid.ColumnSpan = '5' Text = 'Insert WQL query here' TextWrapping = 'Wrap' AcceptsReturn = 'True'
            MinLines = '6'/>
        <ComboBox x:Name = 'LanguageComboBox' Grid.Row = '1' Grid.Column = '4' MaxWidth = '100' IsReadOnly = 'True' SelectedIndex = '0'>
            <TextBlock> WQL </TextBlock>
        </ComboBox>                                           
        <Grid Grid.Row = '4' Grid.Column = '2' ShowGridLines = 'False'>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height = '5'/>       
                <RowDefinition Height = 'Auto'/>  
            </Grid.RowDefinitions>   
            <Button x:Name = 'CreateButton' Content ='Create' Grid.Column = '0' Grid.Row = '1' MaxWidth = '125' IsDefault = 'True'/>       
            <Button x:Name = 'CancelButton' Content = 'Cancel' Grid.Column = '2' Grid.Row = '1' MaxWidth = '125' IsCancel = 'True'/>       
        </Grid>
        <TextBox x:Name = 'StatusTextbox' Grid.Row = '5' Grid.ColumnSpan = '5' IsReadOnly = 'True'>       
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

$reader=(New-Object System.Xml.XmlNodeReader $filter_xaml)
$Global:Filter_Window=[Windows.Markup.XamlReader]::Load( $reader )

##Connect To Controls
$CreateButton = $Filter_Window.FindName('CreateButton')
$CancelButton = $Filter_Window.FindName('CancelButton')
$LanguageComboBox = $Filter_Window.FindName('LanguageComboBox')
$Name_txtbx = $Filter_Window.FindName('Name_txtbx')
$Namespace_txtbx = $Filter_Window.FindName('Namespace_txtbx')
$Query_txtbx = $Filter_Window.FindName('Query_txtbx')
$StatusTextbox = $Filter_Window.FindName('StatusTextbox')

##Events
#Cancel
$CancelButton.Add_Click({    
    $Filter_Window.Close()      
})

#Create
$CreateButton.Add_Click({
    $instanceFilter = ([WMICLASS]"\\$Computername\root\subscription:__EventFilter").CreateInstance()
    $instanceFilter.QueryLanguage = $LanguageComboBox.Text
    $instanceFilter.Query = $Query_txtbx.Text
    $instanceFilter.Name = $Name_txtbx.Text
    $instanceFilter.EventNamespace = $Namespace_txtbx.Text
    Try {
        $result = $instanceFilter.Put()    
        #Create the WMI Event Filter
        $Filter_Window.DialogResult = $True
        $Filter_Window.Close()
    } Catch {
        $StatusTextbox.Foreground = 'Red'
        $StatusTextbox.Text = ("{0}" -f $_.Exception.Message)
    }
})

$Filter_Window.ShowDialog()