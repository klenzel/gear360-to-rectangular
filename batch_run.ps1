$strCurrentLocation = (Get-Location).Path

#Konfiguration
$intBildBreiteInPixel = 7776
$intBildHoeheInPixel = 3888
$strPTGuiExecutable = "C:\Program Files\PTGui\PTGui.exe"
$strIrfanExecutable = $strCurrentLocation + "\irfanview\i_view64.exe"
$strExifToolExecutable = $strCurrentLocation + "\exiftool\exiftool.exe"

#Funktionen
function Start-GPSAbfrage {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $MyForm = New-Object System.Windows.Forms.Form
    $MyForm.Text="GPS-Koordinaten"
    $MyForm.Size = New-Object System.Drawing.Size(320,200)
     
 
    $mLabel1 = New-Object System.Windows.Forms.Label
        $mLabel1.Text="Bitte Koordinaten eingeben."
        $mLabel1.Top="25"
        $mLabel1.Left="13"
        $mLabel1.Anchor="Left,Top"
    $mLabel1.Size = New-Object System.Drawing.Size(300,23)
    $MyForm.Controls.Add($mLabel1)
         
 
    $mLabel2 = New-Object System.Windows.Forms.Label
        $mLabel2.Text="Beispiel: 50.601024, 6.951774"
        $mLabel2.Top="50"
        $mLabel2.Left="13"
        $mLabel2.Anchor="Left,Top"
    $mLabel2.Size = New-Object System.Drawing.Size(250,23) 
    $MyForm.Controls.Add($mLabel2)
         
 
    $mLabel3 = New-Object System.Windows.Forms.Label
        $mLabel3.Text="Eingabe:"
        $mLabel3.Top="79"
        $mLabel3.Left="9"
        $mLabel3.Anchor="Left,Top"
    $mLabel3.Size = New-Object System.Drawing.Size(60,23)
    $MyForm.Controls.Add($mLabel3)
         
 
    $mTextBox2 = New-Object System.Windows.Forms.TextBox
        #$mTextBox2.Text="long"
        $mTextBox2.Top="75"
        $mTextBox2.Left="78"
        $mTextBox2.Anchor="Left,Top"
    $mTextBox2.Size = New-Object System.Drawing.Size(150,23)
    $MyForm.Controls.Add($mTextBox2)
         
 
    $mButton1 = New-Object System.Windows.Forms.Button
        $mButton1.Text="Speichern"
        $mButton1.Top="115"
        $mButton1.Left="20"
        $mButton1.Anchor="Left,Top"
    $mButton1.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $mButton1.Size = New-Object System.Drawing.Size(100,23)
    $MyForm.Controls.Add($mButton1)
         
 
    $mButton2 = New-Object System.Windows.Forms.Button
        $mButton2.Text="Abbrechen"
        $mButton2.Top="115"
        $mButton2.Left="129"
        $mButton2.Anchor="Left,Top"
    $mButton2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $mButton2.Size = New-Object System.Drawing.Size(100,23)
    $MyForm.Controls.Add($mButton2)

    $strErgebnis=$MyForm.ShowDialog()

    if ($strErgebnis -eq "OK") {
        $strEingabe = ($mTextBox2.Text).Replace(" ","")
        $strLat=($strEingabe.Split(","))[0]
        $strLon=($strEingabe.Split(","))[1]
        return "-exif:gpslatitude=" + $strLat + "@-exif:gpslongitude=" + $strLon
    }
    else {
        return $false
    }
}

#Programmablauf
$intHalbeBildBreiteInPixel = $intBildBreiteInPixel / 2
$arrSourceImages = Get-ChildItem "00-Sources"


foreach ($strSourceImage in $arrSourceImages) {
    $strSourceImageName = ($strSourceImage.Name).Replace(".JPG", "")



    #IrfanView => Bild in zwei Hälften teilen
    $strSplitLeftImageName = $strCurrentLocation + "\01-Split\" + $strSourceImageName + "_LEFT.jpg"
    $strSplitRightImageName = $strCurrentLocation + "\01-Split\" + $strSourceImageName + "_RIGHT.jpg"
    
    $strIrfanSource = "`"" + $strCurrentLocation + "\00-Sources\" + $strSourceImageName + ".JPG`""
    $strIrfanToDoA1 = "/crop=(0,0," + $intHalbeBildBreiteInPixel + "," + $intBildHoeheInPixel  +")"
    $strIrfanToDoA2 = "/convert=`"" + $strSplitLeftImageName + "`""
    $strIrfanToDoB1 = "/crop=(" + $intHalbeBildBreiteInPixel + ",0," + $intHalbeBildBreiteInPixel + "," + $intBildHoeheInPixel + ")"
    $strIrfanToDoB2 = "/convert=`"" + $strSplitRightImageName + "`""

    & $strIrfanExecutable $strIrfanSource $strIrfanToDoA1 $strIrfanToDoA2
    & $strIrfanExecutable $strIrfanSource $strIrfanToDoB1 $strIrfanToDoB2



    #PTGui => Dynamische Konfiguration erstellen
    $strDynamicConfigFile = "$strCurrentLocation\02-PTGuiConfig\$strSourceImageName.pts"
    $strOutputFileName = $strCurrentLocation + "\99-Results\" + $strSourceImageName + "_panorama.jpg"

    $strKonfigurationsDatei = Get-Content "$strCurrentLocation\inc\360_template.pts"
    $strKonfigurationsDatei = $strKonfigurationsDatei -replace "@@OUTPUTFILE@@", $strOutputFileName
    $strKonfigurationsDatei = $strKonfigurationsDatei -replace "@@IMAGE_LEFT@@", $strSplitLeftImageName
    $strKonfigurationsDatei = $strKonfigurationsDatei -replace "@@IMAGE_RIGHT@@", $strSplitRightImageName

    Set-Content $strDynamicConfigFile $strKonfigurationsDatei



    #PTGui => Bilder stitchen
    $strPTGUIArg1 = "-batch"
    $strPTGUIArg2 = "-x"

    & $strPTGuiExecutable $strPTGUIArg1 $strPTGUIArg2 $strDynamicConfigFile


    #GPS-Koordinaten setzen
    $strGPSKoordinaten = Start-GPSAbfrage -strBildName $strSourceImage

    if ($strGPSKoordinaten) {
        $strExifToolArg1 = "-overwrite_original"
        $strExifToolArg2 = "-P"
        $strExifToolArg3 = "-q"
        $strExifToolArg4 = ($strGPSKoordinaten.Split("@"))[0]
        $strExifToolArg5 = ($strGPSKoordinaten.Split("@"))[1]
        $strExifToolArg6 = "-exif:gpslatituderef=N"
        $strExifToolArg7 = "-exif:gpslongituderef=E"
        & $strExifToolExecutable $strExifToolArg1 $strExifToolArg2 $strExifToolArg3 $strExifToolArg4 $strExifToolArg5 $strExifToolArg6 $strExifToolArg7 $strOutputFileName
    }


    #Aufräumen
    sleep 5
    Move-Item "00-sources\$strSourceImage" "98-SourceDone\"
    Remove-Item "01-Split\*.jpg"
    Remove-Item "02-PTGuiConfig\*.pts"
}