<#
Author: Michael Shirazi
Date: 1/25/2022
     
  .Synopsis:
    This Windows application loads a dashboard to assist with finding site specific information
  .Description:
      This program works by importing multiple CSVs from NOC's shared directory.
      The imported CSVs contain the circuit information for every location. A different script pulls the CSVs from Google Sheets on a recurring basis.
      The script loads a WinForms GUI which allows you to input a location number which parses through the CSVs to present you with site specific information.
#>


#This imports the CSVs and replaces their headers with numbers to avoid duplicate header name conflicts
$master = Import-Csv '<CSV Filepath>' -Header @(1..36)
$ISP = Import-Csv '<CSV Filepath>' -Header @(1..46)
$Lum = Import-Csv '<CSV Filepath>' -Header @(1..46)
$numbers = Import-Csv '<CSV Filepath>' -Header @(1..10)

#This function parses through the CSVs and presents them to the GUI
function get-data {
  param ()

  $textbox1.Text            = ""
  $TextBox2.Text            = ""
  $textbox3.Text            = ""
  $TextBox6.Text            = ""
  $textbox4.Text            = ""
  $TextBox5.Text            = ""
  $TextBox7.Text            = ""
  $TextBox8.Text            = ""
  $TextBox9.Text            = ""
  $TextBox10.Text           = ""
  $TextBox11.Text           = ""  
  $TextBox99.text           = ""

  
  foreach($m in $master)
  {
   if($m.3 -eq $SearchTextBox.Text -and $m.1 -eq 'Active' ){
  
   $textbox1.Text            = Write-Output $m.4
   $TextBox2.Text            = Write-Output $m.6
   $textbox3.Text            = Write-Output $m.5
   $TextBox6.Text            = Write-Output $m.19
   $textbox4.Text            = Write-Output $m.7
   $TextBox5.Text            = Write-Output $m.8
   
   
   } 
  else {} }
 

   foreach($s in $ISP){

     if($s.3 -eq $SearchTextBox.Text){
    
    $TextBox8.Text           = Write-Output $s.6
    $LecAccountLabel.text    = "LEC Account #"
    $TextBox9.text           = write-output $s.2
    $TextBox7.Text           = Write-Output $s.7

     }
   
  else {} }


  foreach($L in $Lum){

         if ($L.2 -match $SearchTextBox.Text -or $l.1 -match $SearchTextBox.text)
           {
             if ($SearchTextBox.Text -notmatch '\d\d\d\d') {}
             else {
                   
            $TextBox10.Text           = Write-output $L.27
            $textbox1.Text            = Write-Output $L.3 
            $TextBox2.Text            = Write-Output $L.18
            $textbox3.Text            = Write-Output $L.16
            $TextBox6.Text            = ""
            $textbox4.Text            = Write-Output $L.19
            $TextBox5.Text            = Write-Output $L.20
            $TextBox8.Text            = Write-Output ""
            $LecAccountLabel.text     = "LEC Circuit ID"
            $TextBox9.text            = write-output $L.29
            $TextBox7.Text            = Write-Output $L.28
            $TextBox99.text           = write-output $L.32
           }
                    }

        
       foreach($N in $numbers){

             if ($SearchTextBox.Text -eq $n.1) 
                {
                $TextBox11.Text   = Write-Output $n.9
                }
       }     
       
       if (!(test-path '<CSV Filepath>') {$errortxt.text  = "<CSV Filepath> is inaccessible"}
       else {$errortxt.text  = ""}
      } 



}

#### The following code builds the WinForm GUI ####

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(732,575)
$Form.text                       = "NOC Dashboard"
$Form.TopMost                    = $false

$FormTabControl = New-object System.Windows.Forms.TabControl 
$FormTabControl.Size = "372,252" 
$FormTabControl.Location = "319,156" 
$FormTabControl.Width = '375'
$FormTabControl.Height = '371'
$form.Controls.Add($FormTabControl)

$Tab1 = New-object System.Windows.Forms.Tabpage
$Tab1.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab1.UseVisualStyleBackColor = $True 
$Tab1.Name = "Tab1" 
$Tab1.Text = "Phone Numbers” 
$FormTabControl.Controls.Add($Tab1)

$DataGridView3                   = New-Object system.Windows.Forms.DataGridView
$DataGridView3.width             = 520
$DataGridView3.height            = 375
$DataGridView3Data = @(@("Hidden","Hidden"))  
$DataGridView3.ColumnCount = 2
$DataGridView3.ColumnHeadersVisible = $true
$DataGridView3.Columns[0].Name = "ID"
$DataGridView3.Columns[0].Width = 120
$DataGridView3.Columns[1].Name = "Phone Number"
$DataGridView3.Columns[1].Width = 400
foreach ($rows in $DataGridView3Data){
          $DataGridView3.Rows.Add($rows) | Out-Null
      } 
   
$Tab1.Controls.AddRange($DataGridView3)


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.UseVisualStyleBackColor = $True 
$Tab2.Name = "Tab2" 
$Tab2.Text = "VLAN” 
$FormTabControl.Controls.Add($Tab2)

$DataGridView2                   = New-Object system.Windows.Forms.DataGridView
$DataGridView2.width             = 520
$DataGridView2.height            = 375
$DataGridView2Data = @(@("Hidden"))
$DataGridView2.ColumnCount = 2
$DataGridView2.ColumnHeadersVisible = $true
$DataGridView2.Columns[0].Name = "VLAN"
$DataGridView2.Columns[1].Name = "Device Type"
$DataGridView2.Columns[1].Width = 400
foreach ($row in $DataGridView2Data){
          $DataGridView2.Rows.Add($row) | Out-Null
          }
      
$DataGridView2.location          = New-Object System.Drawing.Point(1,2)

$Tab2.Controls.Add($DataGridView2)

$Tab3 = New-object System.Windows.Forms.Tabpage
$Tab3.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab3.UseVisualStyleBackColor = $True 
$Tab3.Name = "Tab3" 
$Tab3.Text = "Demarc” 
$FormTabControl.Controls.Add($Tab3)

$TextBox99                        = New-Object system.Windows.Forms.TextBox
$TextBox99.multiline              = $true
$TextBox99.width                  = 370
$TextBox99.height                 = 400
$TextBox99.location               = New-Object System.Drawing.Point(0,0)
$TextBox99.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox99.text                   = ""

$tab3.Controls.Add($TextBox99)


$TopLabel                        = New-Object system.Windows.Forms.Label
$TopLabel.text                   = "NOC Dashboard"
$TopLabel.AutoSize               = $true
$TopLabel.width                  = 25
$TopLabel.height                 = 10
$TopLabel.location               = New-Object System.Drawing.Point(319,41)
$TopLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$SearchTextBox                   = New-Object system.Windows.Forms.TextBox
$SearchTextBox.multiline         = $false
$SearchTextBox.width             = 79
$SearchTextBox.height            = 20
$SearchTextBox.location          = New-Object System.Drawing.Point(21,103)
$SearchTextBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$SearchTextBox.Add_KeyDown({
              if ($_.KeyCode -eq "Enter") {
                $SearchTextBox.multiline         = $True
                get-data 
                $SearchTextBox.multiline         = $false
               
              }
              })


$LocationNumber                  = New-Object system.Windows.Forms.Label
$LocationNumber.text             = "Location #"
$LocationNumber.AutoSize         = $true
$LocationNumber.width            = 25
$LocationNumber.height           = 10
$LocationNumber.location         = New-Object System.Drawing.Point(26,80)
$LocationNumber.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 371
$Panel1.width                    = 265
$Panel1.location                 = New-Object System.Drawing.Point(21,156)
$Panel1.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#dad9d9")

$SearchButton                    = New-Object system.Windows.Forms.Button
$SearchButton.text               = "Search"
$SearchButton.width              = 52
$SearchButton.height             = 22
$SearchButton.location           = New-Object System.Drawing.Point(105,101)
$SearchButton.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',9)
$SearchButton.add_click({get-data})

$LocationName                    = New-Object system.Windows.Forms.Label
$LocationName.text               = "Location Name"
$LocationName.AutoSize           = $true
$LocationName.width              = 25
$LocationName.height             = 10
$LocationName.location           = New-Object System.Drawing.Point(185,78)
$LocationName.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$errortxt                    = New-Object system.Windows.Forms.Label
$errortxt.text               = ""
$errortxt.AutoSize           = $true
$errortxt.width              = 25
$errortxt.height             = 10
$errortxt.location           = New-Object System.Drawing.Point(10,10)
$errortxt.forecolor          = "Red"


$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 126
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(180,103)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox1.text                   = ""

$CityLabel                       = New-Object system.Windows.Forms.Label
$CityLabel.text                  = "City"
$CityLabel.AutoSize              = $true
$CityLabel.width                 = 25
$CityLabel.height                = 10
$CityLabel.location              = New-Object System.Drawing.Point(324,78)
$CityLabel.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 100
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(319,103)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox2.Text                   = ""

$AddressLabel                    = New-Object system.Windows.Forms.Label
$AddressLabel.text               = "Address"
$AddressLabel.AutoSize           = $true
$AddressLabel.width              = 25
$AddressLabel.height             = 10
$AddressLabel.location           = New-Object System.Drawing.Point(443,80)
$AddressLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 119
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(436,103)
$TextBox3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$textbox3.text                   = ""

$StateLabel                      = New-Object system.Windows.Forms.Label
$StateLabel.text                 = "State"
$StateLabel.AutoSize             = $true
$StateLabel.width                = 25
$StateLabel.height               = 10
$StateLabel.location             = New-Object System.Drawing.Point(576,80)
$StateLabel.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox4                        = New-Object system.Windows.Forms.TextBox
$TextBox4.multiline              = $false
$TextBox4.width                  = 41
$TextBox4.height                 = 20
$TextBox4.location               = New-Object System.Drawing.Point(573,103)
$TextBox4.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ZipLabel                        = New-Object system.Windows.Forms.Label
$ZipLabel.text                   = "ZIP"
$ZipLabel.AutoSize               = $true
$ZipLabel.width                  = 25
$ZipLabel.height                 = 10
$ZipLabel.location               = New-Object System.Drawing.Point(637,80)
$ZipLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox5                        = New-Object system.Windows.Forms.TextBox
$TextBox5.multiline              = $false
$TextBox5.width                  = 55
$TextBox5.height                 = 20
$TextBox5.location               = New-Object System.Drawing.Point(632,103)
$TextBox5.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox6                        = New-Object system.Windows.Forms.TextBox
$TextBox6.multiline              = $false
$TextBox6.width                  = 130
$TextBox6.height                 = 20
$TextBox6.location               = New-Object System.Drawing.Point(23,32)
$TextBox6.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ATTCitcuitLabel                 = New-Object system.Windows.Forms.Label
$ATTCitcuitLabel.text            = "AT&T Circuit ID"
$ATTCitcuitLabel.AutoSize        = $true
$ATTCitcuitLabel.width           = 25
$ATTCitcuitLabel.height          = 10
$ATTCitcuitLabel.location        = New-Object System.Drawing.Point(28,8)
$ATTCitcuitLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$LECLabel                        = New-Object system.Windows.Forms.Label
$LECLabel.text                   = "LEC"
$LECLabel.AutoSize               = $true
$LECLabel.width                  = 25
$LECLabel.height                 = 10
$LECLabel.location               = New-Object System.Drawing.Point(28,115)
$LECLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox7                        = New-Object system.Windows.Forms.TextBox
$TextBox7.multiline              = $false
$TextBox7.width                  = 130
$TextBox7.height                 = 20
$TextBox7.location               = New-Object System.Drawing.Point(24,139)
$TextBox7.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox7.Text                   = ""

$ProviderLabel                   = New-Object system.Windows.Forms.Label
$ProviderLabel.text              = "Provider"
$ProviderLabel.AutoSize          = $true
$ProviderLabel.width             = 25
$ProviderLabel.height            = 10
$ProviderLabel.location          = New-Object System.Drawing.Point(31,62)
$ProviderLabel.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))


$TextBox8                        = New-Object system.Windows.Forms.TextBox
$TextBox8.multiline              = $false
$TextBox8.width                  = 130
$TextBox8.height                 = 20
$TextBox8.location               = New-Object System.Drawing.Point(24,85)
$TextBox8.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox8.Text                   = "" 


$LecAccountLabel                 = New-Object system.Windows.Forms.Label
$LecAccountLabel.text            = "LEC Account #"
$LecAccountLabel.AutoSize        = $true
$LecAccountLabel.width           = 25
$LecAccountLabel.height          = 10
$LecAccountLabel.location        = New-Object System.Drawing.Point(28,175)
$LecAccountLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox9                        = New-Object system.Windows.Forms.TextBox
$TextBox9.multiline              = $false
$TextBox9.width                  = 129
$TextBox9.height                 = 20
$TextBox9.location               = New-Object System.Drawing.Point(23,201)
$TextBox9.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox9.Text                   = ""

$LumenCircuitLabel               = New-Object system.Windows.Forms.Label
$LumenCircuitLabel.text          = "Lumen Circuit ID"
$LumenCircuitLabel.AutoSize      = $true
$LumenCircuitLabel.width         = 25
$LumenCircuitLabel.height        = 10
$LumenCircuitLabel.location      = New-Object System.Drawing.Point(28,237)
$LumenCircuitLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox10                       = New-Object system.Windows.Forms.TextBox
$TextBox10.multiline             = $false
$TextBox10.width                 = 132
$TextBox10.height                = 20
$TextBox10.location              = New-Object System.Drawing.Point(23,262)
$TextBox10.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox10.Text                  = ""

$LocationNum                     = New-Object system.Windows.Forms.Label
$LocationNum.text                = "Site Phone #"
$LocationNum.AutoSize            = $true
$LocationNum.width               = 25
$LocationNum.height              = 10
$LocationNum.location            = New-Object System.Drawing.Point(26,301)
$LocationNum.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',11,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TextBox11                       = New-Object system.Windows.Forms.TextBox
$TextBox11.multiline             = $false
$TextBox11.width                 = 133
$TextBox11.height                = 20
$TextBox11.location              = New-Object System.Drawing.Point(23,324)
$TextBox11.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox11.Text                  = ""

#This loads all of the winform elements into the form
$Form.controls.AddRange(@($TopLabel,$SearchTextBox,$LocationNumber,$Panel1,$SearchButton,$LocationName,$TextBox1,$CityLabel,$TextBox2,$AddressLabel,$TextBox3,$StateLabel,$TextBox4,$ZipLabel,$TextBox5,$errortxt))
$Panel1.controls.AddRange(@($TextBox6,$ATTCitcuitLabel,$LECLabel,$TextBox7,$ProviderLabel,$TextBox8,$LecAccountLabel,$TextBox9,$LumenCircuitLabel,$TextBox10, $LocationNum, $TextBox11))



[void]$Form.ShowDialog()