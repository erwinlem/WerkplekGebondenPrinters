[cmdletbinding()]
param(
      [string]
      $filterDescription = "^(ETK-Medicatie|ETK-LAB|ETK-RODEBALK|ETK-STAM|A4|ETK-Patient|Polsband|PolsbandB|PolsbandK)$"
)

Import-Module activedirectory
Add-Type -AssemblyName System.Drawing
Add-Type -Assembly PresentationFramework


write-verbose("starting, $PSScriptRoot");

# we moeten moeilijk doen omdat wpf anders de files lockt
function ConvertTo-BitmapImage {
    param([
        Parameter(ValueFromPipeline = $true)]
        [string[]]$base64
    )

    process {
        foreach ($str in $base64) {
            $bmp = [System.Drawing.Bitmap]::FromFile($str)

            $memory = New-Object System.IO.MemoryStream
            $null = $bmp.Save($memory, [System.Drawing.Imaging.ImageFormat]::Png)
            $memory.Position = 0

            $img = New-Object System.Windows.Media.Imaging.BitmapImage
            $img.BeginInit()
            $img.StreamSource = $memory
            $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $img.EndInit()
            $img.Freeze()

            $memory.Close()
            $bmp.Dispose()
            $img
        }
    }
}

function waitcursor([bool] $a) {
    # kan geen windows cursor dingen doen in class, superstom
    if ($a) {
        return [System.Windows.Input.Cursors]::Wait;
    } else {
        return [System.Windows.Input.Cursors]::Arrow;    
    }

}

# Define a .NET class representing your data object
class WindowData {
    [scriptblock] $butt_press = { 
        write-verbose ("button press "+$args.Name);
        switch ($args.Name) {
            "butt_reset" {
                write-verbose ("resetting")
                refreshPrinters
                refresh
                refreshType
                write-verbose ("klaar met resetting")
            }
            'butt_cancel' {
                $Form.Close();        
            }   
            'butt_set' {
                $r = $global:Datatable_type.Rows |? { $_.Type -eq $Form.FindName("printers").selecteditem.description};
                $r.Printer = $Form.FindName("printers").selecteditem.printername;
                $r.Locatie = $Form.FindName("printers").selecteditem.location;
            }
            'butt_unset' {
                $r = $Form.FindName("printerType").selecteditem;
                $r.Printer = "";
                $r.Locatie = "";

            }
            'butt_ok' {
                
                $Form.Cursor =  waitcursor($true);
                $printersNew = @($Form.FindName("printerType").Items |? { $_.Printer -notmatch "^$"} |% { $_.Printer });
                $printersOld = @($printersCurrent |% { $_.sharename });

                write-verbose("nieuwe printers : " + ($printersNew -join ",") );
                write-verbose("oude printers : " + ( $printersOld -join ",") );

                Compare-Object $printersNew $printersOld |? { $_.InputObject -ne $null }|% {

                    # geen start-job gebruiken hiervoor : https://github.com/PowerShell/PowerShell/issues/11659
                    $printer=  $_.InputObject;
                    $printer = $printers |? { $_.printername -eq $printer } |% { $_.uncname -replace '.intranet.local', ''}

                    write-verbose("printer " + ($_) );
                    switch ($_.SideIndicator) {
                        "==" {
                            write-verbose "keeping "+$printer;
                        }
                        "<=" {
                            write-verbose("printer $printer toevoegen" );
                            Add-Printer -ConnectionName $printer;
                            write-verbose("printer $printer toegevoegd" );
                        }
                        "=>" {
                            write-verbose("weg printer " + ($printer) );
                            get-printer |? { $_.name -eq $printer } | Remove-Printer
                            write-verbose("weggehaald printer " + ($printer) );
                        }
                        default {
                            write-verbose ("Geen idee wat dit is "+$printer);
                        }

                    }
                }

                refreshPrinters
                $Form.Cursor = waitcursor($false);
    
                $printersNew |% {
                    $printer = $_;
                    $printers |? { $_.printername -eq $printer } |% { $_.uncname -replace '.intranet.local', ''}  
                } | sort-object | out-file ("$PSScriptRoot\computers\"+$env:CLIENTNAME+".txt")
                write-verbose "KLAAR!"
            }

        }  
        
    };
    [string] $TB_Werkplek = "Werkplek - {0}" -f @( $env:CLIENTNAME);

    ViewModel() {
     #   $this.SayHello = { [System.Windows.MessageBox]::Show("Hello from DataContext!") }
    }

}

$WindowData = [WindowData]::new()

$global:selected_type = ""

# voor get-adobject is de activedirectory module nodig
$printers = Get-AdObject -filter "objectCategory -eq 'printqueue'" -Properties servername, printername, portname, description,location,uNCName |? { $_.Description -match $filterDescription } | sort-object printername;

function refreshPrinters() {
    $global:printersCurrent = get-printer | select sharename,comment, location |? { $_.comment -match $filterDescription };
    $global:printerType = ($printers | group-object -Property description | select name) |% {
        $t = $_.name;
        [pscustomobject]@{
            "Type"=$t;
            "Printer"= (($printersCurrent |? { $_.comment -eq $t}) |% { $_.sharename } | select-object -first 1);
            "Locatie"= (($printersCurrent |? { $_.comment -eq $t}) |% { $_.location }  | select-object -first 1);
        }
    } | sort-object -Property Type;

    $printerType | out-string | write-verbose;
}

refreshPrinters

# Als je een editor nodig hebt : https://xaml.io/
[xml]$XAML = @"
<Window Title="Werkplekgebondenprinter 2.0" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
      Foreground="{DynamicResource Theme_TextBrush}">

<Page>
   <Grid  Background="{DynamicResource Theme_BackgroundBrush}">
        <TextBlock Name="TB_Werkplek" Text="{Binding TB_Werkplek}" FontSize="20" Margin="8,8,0,8" HorizontalAlignment="Left" VerticalAlignment="Top" Height="32" Width="624" />
        <TextBlock Text="Filter:" FontSize="14" Margin="648,16,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Height="24" />
        <TextBox Name="TB_Filter" Margin="720,8,4,8" FontSize="24" VerticalAlignment="Top" Height="32" IsEnabled="True"  />
        <Button Name="butt_set" Content="← GEBRUIK" Width="70" Height="40" Margin="640,108,4,4" VerticalAlignment="Top" HorizontalAlignment="Left"   />
        <Button Name="butt_reset" Command="{Binding SayHello}" Content="↺ Herstel" Width="70" Height="40" Margin="640,156,4,4" VerticalAlignment="Top" HorizontalAlignment="Left"   />
        <Button Name="butt_unset" Content="❌ Verwijder" Width="70" Height="40" Margin="640,206,4,4" VerticalAlignment="Top" HorizontalAlignment="Left"   />
        <DataGrid Name="printerType" ItemsSource="{Binding SampleData.Employees}" Margin="8,48,0,328" HorizontalAlignment="Left" Width="624" IsReadOnly="True" SelectionMode="Single" CanUserSortColumns="True"  />
        <DataGrid Name="printers"  ItemsSource="{Binding SampleData.Employees}" Margin="720,48,8,48"  IsReadOnly="True" SelectionMode="Single"  />
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="4,4,4,4">
            <Button Name="butt_cancel" Content="Annuleren" Width="100" Margin="4,4,4,4"  />
            <Button Name="butt_ok" Content="Toepassen" Margin="4,4,4,4" Width="100" />
        </StackPanel>
          <Image Name="img_voorbeeld" Margin="8,0,0,8" HorizontalAlignment="Left" VerticalAlignment="Bottom" Width="320" Height="320" />
    </Grid>
</Page>
</Window> 
"@ 

$Form=[Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $XAML) )
$Form.DataContext = $WindowData

# knoppen verbinden met class (powershell doet niet aan ICommand)
"set unset ok reset cancel" -split " " |% {
    $Form.FindName("butt_$($_)").add_Click($WindowData.butt_press)
}

# filter verbinden
$Form.FindName('TB_Filter').add_TextChanged({
    param([System.Windows.Controls.TextBox]$sender, $e)
    refresh    
})


$Form.FindName('img_voorbeeld').Source = (ConvertTo-BitmapImage "$($PSScriptRoot)\voorbeeld\kat.png")

$Form.FindName('printerType').AddHandler([System.Windows.Controls.DataGrid]::SelectionChangedEvent, [System.Windows.RoutedEventHandler] {
    param([System.Windows.Controls.DataGrid]$sender, $e)
    $global:selected_type = $e.AddedItems[0].type;
    $img = "$($PSScriptRoot)\voorbeeld\$($e.AddedItems[0].type).png";
    if (!(test-path $img)) {
        $img = "$($PSScriptRoot)\voorbeeld\kat.png" # indien niet gevonden een leuk kattenplaatje.
    }
    $Form.FindName('img_voorbeeld').Source = (ConvertTo-BitmapImage $img);
    refresh;
}
);

function refreshType() {
    $global:Datatable_type = New-Object System.Data.DataTable
    [void]$Datatable_type.Columns.AddRange("Type Printer Locatie" -split " ");
    $printerType |% {
        [void]$Datatable_type.Rows.Add($_.Type, $_.Printer, $_.Locatie)
    }
    $Form.FindName("printerType").ItemsSource = $Datatable_type.DefaultView;
}

refreshType

$Datatable = New-Object System.Data.DataTable
[void]$Datatable.Columns.AddRange("printername description location" -split " ");

function refresh() {
    $InputText = $Form.FindName('TB_Filter').text;
    $s = $global:selected_type;
    $filter = "(printername LIKE '%$InputText%' or location LIKE '%$InputText%')";
    if ($s -and $s -ne "") {
        $filter += " and description like '$s'";
    }
    $Datatable.DefaultView.RowFilter = $filter
}

# setup detail view
$printers |% {
    [void]$Datatable.Rows.Add( ($_.printername, $_.description, $_.location ));
}

$Form.FindName("printers").ItemsSource = $Datatable.DefaultView
$Form.ShowDialog()

write-verbose("end of script");