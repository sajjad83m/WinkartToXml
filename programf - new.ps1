
Write-host -b Black -f Green ("-"*([Console]::WindowWidth-1))
Write-Host -b Black -f Green "`n`n"
$a = " Convert Output Winkart's Xls files to a Xml file "
Write-Host -b Black -f Green "$("-"*(([Console]::WindowWidth-1-$a.Length)/2))$a$("-"*(([Console]::WindowWidth-1-$a.Length)/2))"
Write-Host -b Black -f Green "`n`n"
$b = " by Sajjad (@Sajjad_moh) "
Write-Host -b Black -f Green "$("-"*(([Console]::WindowWidth-1-$b.Length)/2))$b$("-"*(([Console]::WindowWidth-1-$b.Length)/2))"
Write-Host -b Black -f Green "`n`n"
Write-host -b Black -f Green ("-"*([Console]::WindowWidth-1))
$host.UI.RawUI.BackgroundColor="black"

function ExitProgram(){
         
   write-Host -b Black -NoNewline "press 0 to exit: "
   do{
       
       $key = $host.UI.RawUI.ReadKey().Character
   }until($key-eq '0')
   exit

}

function Show-ProgressBar {
     param (
         [string]$Caption,
         [int]$PercentComplete
         
     )

     $ProgressBarLength = 40
     $CompletedLength = [math]::Round(($PercentComplete / 100) * $ProgressBarLength)
     $RemainingLength = $ProgressBarLength - $CompletedLength

     $ProgressBar = ('■' * $CompletedLength) + ('-' * $RemainingLength)
     Write-Host -NoNewline -f Green -b Black "`r$caption`: [$ProgressBar] $($PercentComplete.ToString("0.00"))%"
     
 }

function Convert-ShamsiToGregorian($Date){

    $dates = $Date -split '/'
    $persianCalendar = New-Object System.Globalization.PersianCalendar
    try{
        $gregorianDate = $persianCalendar.ToDateTime($dates[0], $dates[1], $dates[2], 0, 0, 0, 0)
        $date2 = $gregorianDate.ToString("yyyy/MM/dd")
        return $date2
    }catch{
        return $null
    }
}
        
function Convert-GregorianToShamsi($Date){

    $gregorianDate = Get-Date $Date
    $shamsiCalendar = New-Object System.Globalization.PersianCalendar
    $shamsiYear = $shamsiCalendar.GetYear($gregorianDate)
    $shamsiMonth = $shamsiCalendar.GetMonth($gregorianDate)
    $shamsiDay = $shamsiCalendar.GetDayOfMonth($gregorianDate)

    $shamsiDate = "{0:0000}/{1:00}/{2:00}" -f $shamsiYear, $shamsiMonth, $shamsiDay

    return $shamsiDate

}

function get-MinMax-date($Dates){

   $minDate = $Dates | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
   $maxDate = $Dates | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

   return @($minDate,$maxDate)
}

function get-Date-Of-file($file){

   $startDate = (($file-split"در تاریخ ")[1]-split"تا تاریخ ")[0]

   $endDate = (($file-split"تا تاریخ ")[1]-split" حضور دارند")[0]

   $startTime = (($file-split'بازه : ')[1]-split' الی ')[0]

   $endTime = ((($file-split'<Data ss:Type="String">')[1]-split'الی ')[1]-split'</Data>')[0]

   return @($startDate,$endDate,$startTime,$endTime)
}

function addDayss($Date,$n){

   [datetime]$gregorianDate = Convert-ShamsiToGregorian $Date
   $newGregorianDate = $gregorianDate.AddDays($n)
   $newShamsiDate = Convert-GregorianToShamsi $newGregorianDate

   return $newShamsiDate
}

function SearchInFile($file,$text){

    if($file.Contains($text)){
        return "TRUE"


    }else{

        return "FALSE"
    }
}

function AddRow($PrimaryXMLFile,$newXMLPath,$index,$name,$family,$number,$present){

  $add = "<Row ss:AutoFitHeight=`"0`" ss:Height=`"17.0625`">
    <Cell ss:StyleID=`"s418`"><Data ss:Type=`"Number`">$index</Data></Cell>
    <Cell ss:StyleID=`"s67`"><Data ss:Type=`"String`">$name</Data></Cell>
    <Cell ss:StyleID=`"s67`"><Data ss:Type=`"String`">$family</Data></Cell>
    <Cell ss:StyleID=`"s418`"><Data ss:Type=`"Number`">$number</Data></Cell>
   </Row>"
  
  $data0 = [System.IO.File]::ReadAllText($PrimaryXMLFile)
  $newdata = $data0.replace(("</Table>"),($add+"`n</Table>"))
  [xml]$xml = $newdata

  [int]$rawNumber = $xml.Workbook.Worksheet.table.ExpandedRowCount
  $newdata=($newdata -split "ExpandedRowCount=`"")[0]+"ExpandedRowCount=`""+[String]($rawNumber+1)+
  "`" x:FullColumns="+($newdata -split "`" x:FullColumns=")[1]

  $newdata | Out-File $newXMLPath
}

function AddCloumn($PrimaryXMLFile,$newXMLPath,$presentDate){
 
  $data0 = [System.IO.File]::ReadAllText($PrimaryXMLFile)

  $pe=0
  #forEach($Date in $presentDate.keys){
  forEach($p in $presentDate){
    $Date = $p.keys
  
    $newdata=""
    [xml]$xml=$data0
    [int]$i = 0
    forEach($d in ($s = $data0 -split "</Row>")){

      if($i -gt 1 -and $i -ne ($s.count-1) -and ($i -ne 1)){
          $number = $xml.Workbook.Worksheet.Table.Row[$i].cell[3].InnerText

          if($p.values.Contains($number)){
              $newdata+=$d+" <Cell ss:StyleID=`"s68`"><Data ss:Type=`"String`">P</Data></Cell>"+"`n      </Row>"
          }else{
              $newdata+=$d+" <Cell ss:StyleID=`"s66`"><Data ss:Type=`"String`">غ</Data></Cell>"+"`n      </Row>"
          }

      } elseif($i -ne ($s.count-1)-and ($i -ne 1)){
          $newdata+=$d+"</Row>"

      } elseif($i -eq ($s.count-1)){
          $newdata+=$d
      } elseif($i -eq 1){
          $year=($Date -split '/')[0]
          $year = $year[1]+$year[2]+$year[3]
          $month=($Date -split '/')[1]
          $day=($Date -split '/')[2]
          $newdata+=$d + " <Cell ss:StyleID=`"s65`"><Data ss:Type=`"String`">$year&#10;/$month&#10;/$day</Data></Cell>"+"`n   </Row>"
      }
      $i++
    }

    $columnNumber = (($newdata -split "ExpandedColumnCount=`"")[1] -split "`" ss:ExpandedRowCount")[0]
    #$columnNumber
    $newdata = ($newdata -split "ExpandedColumnCount=`"")[0]+"ExpandedColumnCount=`""+
    +([int]$columnNumber+1)+"`" ss:ExpandedRowCount"+($newdata -split "`" ss:ExpandedRowCount")[1]
    #$newdata
    [xml]$xml = $newdata
    [int]$across = $xml.Workbook.Worksheet.Table.row[0].cell.MergeAcross
    
    $newdata = ($newdata -split "MergeAcross=`"")[0]+"MergeAcross=`""+[string]($across+1)+
    "`" ss:StyleID=`""+((($newdata -split "MergeAcross=`"")[1]) -split "`" ss:StyleID=`"")[1]
    #$newdata
    $newdata = ($newdata -split "<Column ss:StyleID=`"s62`" ss:Width=`"90`"/>")[0]+
    "<Column ss:StyleID=`"s62`" ss:Width=`"90`"/>"+"`n"+
    "<Column ss:StyleID=`"s62`" ss:AutoFitWidth=`"0`" ss:Width=`"22.5`"/>"+
    (($newdata -split "<Column ss:StyleID=`"s62`" ss:Width=`"90`"/>")[1] -split "<Row ss:AutoFitHeight=`"0`" ss:Height=`"37.5`">")[0]+
    "<Row ss:AutoFitHeight=`"0`" ss:Height=`"37.5`">"+
    ($newdata -split "<Row ss:AutoFitHeight=`"0`" ss:Height=`"37.5`">")[1]
    #$newdata
    $data0 = $newdata
    $pe+=1
    $percentComplete = [math]::Round(($pe / $presentDate.keys.Count) * 100, 2)
    Show-ProgressBar -Caption "AddColumn: add Date $Date to outputFile" -PercentComplete $percentComplete
    }
  $newdata | Out-File $newXMLPath
  Write-Host -b Black ""
  Write-Host -b Black -f Green "process has been Completed!"
}

function get-XLSfilesText(){
    
  $filesArray = @{}
  Get-ChildItem ./Input | Select-Object Name | Where-Object{$_.Name -like "*.xls"} | ForEach-Object{

          $filesArray.Add($_.Name,[System.IO.File]::ReadAllText('./Input/'+$_.Name))| Out-Null
          
  }
  if($filesArray.Count -le 0){
    Write-Host -b Black -f Red "Error: the 'Input' folder is empty please add your Wikart's XLS files to it and again run the program"
    ExitProgram
  }


  [int]$pe=0
  forEach($file in $filesArray.Keys){
  
      $BeginningYear = ((get-Date-Of-file $filesArray.Item($file))[0]-split'/')[0]
      $EndingYear = ((get-Date-Of-file $filesArray.Item($file))[1]-split'/')[0]
  
      if($BeginningYear -ne $EndingYear){
          Write-Host -b Black -f red "Error: in the file `"$file`" the Beginning year should be same with the ending year, but $BeginningYear and $EndingYear are different Years"
            
          write-Host -b Black -NoNewline "press 0 to exit: "
          do{
              
              $key = $host.UI.RawUI.ReadKey().Character
          }until($key-eq '0')
          exit
      }
      $pe+=1
      $percentComplete = [math]::Round(($pe / $filesArray.keys.Count) * 100, 2)
  
      Show-ProgressBar -Caption "check dates of input files" -PercentComplete $percentComplete
  }
  Write-Host -b Black ""
  Write-Host -b Black -f Green "process has been Completed!"


  return $filesArray.Clone()

}

function Add_Dates($configFile){

   $settings = [System.IO.File]::ReadAllText($configFile)
   $Dates0 = ((((($settings -split "Date:")[1] -split '{')[1]-split '}')[0].Replace('\','/').Replace('+=','=+').Replace(' ',''))) -split"`n"
   $Dates1 = New-Object System.Collections.ArrayList
   $Dates2  = New-Object System.Collections.ArrayList

   $filesArray = get-XLSfilesText

   forEach($file in $filesArray.values){

       $EndfileDate = (get-Date-Of-file $file)[1]
       if((Convert-ShamsiToGregorian $biggerDate) -lt (Convert-ShamsiToGregorian $EndfileDate)){
       $biggerDate = $EndfileDate
       }
   }

   forEach($Date0 in $Dates0){
       $Date0 = ($Date0 -split "#")[0]
       $Date0 = $Date0.Trim()
       if($Date0.Length -gt 1){

           $Dates1.Add($Date0) | Out-Null
       }
   }
   
   forEach ($Date1 in $Dates1){

      if(($Date1.Length -eq 25 -or $Date1.Length -eq 13) -and ($Date1.Contains( "=+"))){

          if ($Date1.Contains("<<")-and $Date1.Length -eq 25){
              $endDate = ($Date1 -split "<<")[1]
              #$Dates.Remove($Date)

          }elseif($Date1.Length -eq 13 -and !$Date1.Contains("<<")){
   
              $endDate = $biggerDate
              #$Dates.Remove($Date)
   
          }else{
              Write-Host -b Black -f Red "Warning: the date $Date1 is a wong input date"
          }

          $n = (($Date1 -split "=+")[1]-split"<<")[0]
          $Date1 = ($Date1 -split"=+")[0]

          while([datetime](Convert-ShamsiToGregorian $endDate) -ge [datetime](Convert-ShamsiToGregorian $Date1)){
             #Write-Host "end$(Convert-ShamsiToGregorian $endDate)"
             #Write-Host "$(Convert-ShamsiToGregorian $Date1)"
              $Dates2.Add($Date1) | Out-Null
              $Date1 = addDayss $Date1 $n
          }
      }elseif($Date1.Length -ge 8 -and (Convert-ShamsiToGregorian $Date1) -ne $null){
              
              $Dates2.Add($Date1) |Out-Null
      
      }else{
        Write-Host -b Black -f Red "Warning: the date $Date1 is a wong input date"
      
      }
   
   }
   if($Dates2.Count -le 0){
      Write-Host -b Black -f Red "Error: there no exist any date into 'config.conf' file, please add your dates to 'config.conf' file without using '#' before dates"
      ExitProgram
   }
   return $Dates2
}

function get_informations($Infomation_XMLFile){

[xml]$Infomations = [System.IO.File]::ReadAllText($Infomation_XMLFile)
$Rows = $Infomations.Workbook.Worksheet.Table.Row
$Array_informations = @{'index' = @('name';'family';'number')}
$Array_informations.Clear()
[int]$index = 1
[int]$RowNumber = 1
forEach($Row in $Rows){
    
    $name = $Row.cell[0].InnerText.Trim()
    $family = $Row.cell[1].InnerText.Trim()
    $number = $Row.cell[2].InnerText.Trim()
    $RowNumber+=1
    if($number -as [long] -is [long] -and $name -as [long] -isnot [long] -and $family -as [long] -isnot [long]){
        $Array_informations.Add($index,@($name,$family,$number))
        
        #$name;$family;$number;
        $index+=1
    }else{
        
         switch($true) {

             {($number -as [long] -isnot [long])}{
                 Write-Host -f Red -b Black "ERROR:At Row: $RowNumber Cell:3 in the `"informations.xml`" file the personnel number $number is wrong.`nNOTE-THAT: the personnel code column should be on the left side of the `"informations.xml`" file.`n"


             }{($name -as [long] -is [long])}{
                 Write-Host -f Red -b Black "ERROR:At Row: $RowNumber Cell:1 in the `"informations.xml`" file the name $name is wrong.`nNOTE-THAT: the name column should be on the right side of the `"informations.xml`" file.`n"


             }{($family -as [long] -is [long])}{
                 Write-Host -f Red -b Black "ERROR:At Row: $RowNumber Cell:2 in the `"informations.xml`" file the fammily $family is wrong.`nNOTE-THAT: the family column should be in the center of the `"informations.xml`" file.`n"
             }
         }
    
    }

}
if($Array_informations.Count -le 0){
    Write-Host -b Black -f Red "Error: There no exist informations of anyone please check the 'informations.xml' file and add informations to it"
    ExitProgram
}
return $Array_informations
}

function checkMainFilesExisting(){
 
     $configFile = @'
     #EXAMLPES
     Date:{
     #1401/05/12=+7<<1401/08/03 #یک روز در هر هفته بین 2 تاریخ
     #1401/05/12=+1             #هر روز تا بزرگ ترین تاریخ در فایل های ورودی
     #1401/10/12                #
     #1401/01/01=+2             #یک روز در میان تا بزرگ ترین تاریخ در فایل های ورودی
     
     }
     
     Page-Height:50

     Page-Width:20
     
     Have-Header:yes            #no
'@

    $informationsFile = @'
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>@sajjad_moh</Author>
  <LastAuthor>@sajjad_moh</LastAuthor>
  <Created>2023-10-05T17:30:48Z</Created>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>5985</WindowHeight>
  <WindowWidth>15345</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <MaxChange>0.0001</MaxChange>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1" ss:RightToLeft="1">
  <Table ss:ExpandedColumnCount="3" ss:ExpandedRowCount="1" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="110.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="117.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="98.25"/>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">نام</Data></Cell>
    <Cell><Data ss:Type="String">نام خانوادگی</Data></Cell>
    <Cell><Data ss:Type="String">کد پرسنلی</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Selected/>
   <DisplayRightToLeft/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>7</ActiveRow>
     <ActiveCol>1</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
'@
    $exit = $false
    if(![system.IO.File]::Exists("config.conf")){
         [System.IO.File]::WriteAllText("config.conf",$configFile)
         $exit = $true
         Write-Host -b Black -f Green "successfully added 'config.conf' file, please add dates and set Page-height, Page-Width to it"

    }

    if(![System.IO.File]::Exists("informations.xml")){
        [System.IO.File]::WriteAllText("informations.xml",$informationsFile)
        $exit = $true
        Write-Host -b Black -f Green "successfully added informations.xml file, please add persons' informations to it"
    }
    if(![System.IO.Directory]::Exists("./Output")){
        [System.IO.Directory]::CreateDirectory("./Output") | Out-Null
        Write-Host -b Black -f Green "successfully added Output folder, the final xml files will be placed in this folder"

    }
    if(![System.IO.Directory]::Exists("./Input")){
    [System.IO.Directory]::CreateDirectory("./Input") | Out-Null
        $exit = $true
        Write-Host -b Black -f Green "successfully added `"Input`" folder, please add your Winkart's XLS files in the folder and run again the program"
    }
    if($exit){ExitProgram}
}

checkMainFilesExisting

$filesArray = get-XLSfilesText

$informations = get_informations 'informations.xml'

function Get-heightWidth-andHeader($configFile){

    $File = [System.IO.File]::ReadAllLines($configFile)
    $File.ForEach{if($_ -like '*Page-Height:*'){$Height=(($_-split'Page-Height:')[1].replace(' ','')-split'#')[0].Trim()}}
    $File.ForEach{if($_ -like '*Page-Width:*'){$Width=(($_-split'Page-Width:')[1].replace(' ','')-split'#')[0].Trim()}}
    $File.ForEach{if($_ -like '*Have-Header:*'){$Header=(($_-split'Have-Header:')[1].replace(' ','')-split'#')[0].Trim()}}
    try{
       $Header = $Header.ToLower()
    
    }catch{
    
       <#Write-Host -b Black -f red "Error:could not find the string of `"Have-Header:`" please enter the 'yes' or 'no' front of `"Have-Header:`" in config.conf.`nif you still have this problem, remove `'config.conf`' file and run again this program to add `'config.conf`' file and you can see examples in the file"
         
       write-Host -b Black -NoNewline "press 0 to exit: "
       do{
           
           $key = $host.UI.RawUI.ReadKey().Character
       }until($key-eq '0')
       exit
       #>

    }

    if($Height.Length -le 0 -or
       $Height -as [int] -isnot [int]){
       Write-Host -b Black -f red "Error:could not find the number of `"Page-Height:`" please enter the height number front of `"Page-Height:`" in config.conf.`nif you still have this problem, remove `'config.conf`' file and run again this program to add `'config.conf`' file and you can see examples in the file"
         
       write-Host -b Black -NoNewline "press 0 to exit: "
       do{
           
           $key = $host.UI.RawUI.ReadKey().Character
       }until($key-eq '0')
       exit
    }
        #Write-Host "width:$($Width),$($Width.Length),$($Width -as [int] -is [int]),$($Width -eq $null)"
        #Write-Host $('a' -eq 'a')
    if($Width.Length -le 0 -or
       $Width -as [int] -isnot [int]){
       Write-Host -b Black -f red "Error:could not find the number of `"Page-With:`" please enter the height number front of `"Page-With:`" in config.conf.`nif you still have this problem, remove `'config.conf`' file and run again this program to add `'config.conf`' file and you can see examples in the file"
         
       write-Host -b Black -NoNewline "press 0 to exit: "
       do{
           
           $key = $host.UI.RawUI.ReadKey().Character
       }until($key-eq '0')
       exit
    }
    <#if($Header -ne 'yes' -or $Header -ne 'no'){
       Write-Host -b Black -f red "Error:could not find the string of `"Have-Header:`" please enter the 'yes' or 'no' front of `"Have-Header:`" in config.conf.`nif you still have this problem, remove `'config.conf`' file and run again this program to add `'config.conf`' file and you can see examples in the file"
         
       write-Host -b Black -NoNewline "press 0 to exit: "
       do{
           
           $key = $host.UI.RawUI.ReadKey().Character
       }until($key-eq '0')
       exit
    }
    #>
    #[int]$Height=+3
    return @($Height,$Width,$Header)
}

function outputXMLFile($number,$Date){

    $gregorianDate = $Date
    $shamsiCalendar = New-Object System.Globalization.PersianCalendar
    $shamsiYear = "{0:D2}" -f ($shamsiCalendar.GetYear($gregorianDate)%100)
    $shamsiMonth ="{0:D2}" -f ($shamsiCalendar.GetMonth($gregorianDate))
    $shamsiDay ="{0:D2}" -f ($shamsiCalendar.GetDayOfMonth($gregorianDate))
    $shamsiHour ="{0:D2}" -f ($shamsiCalendar.GetHour($gregorianDate))
    $shamsiMinute ="{0:D2}" -f ($shamsiCalendar.GetMinute($gregorianDate))

    return "./Output/out$number-$shamsiYear-$shamsiMonth-$shamsiDay $shamsiHour`;$shamsiMinute.xml"

}
$DateOutput = Get-Date



function set-primartXMLFile($file,$primaryXMLFile,$DateOutput,$HaveHeader){

    $BeginningTime = (get-Date-Of-file $file)[2]
    $EndTime = (get-Date-Of-file $file)[3]
    $BeginningDate = (get-Date-Of-file $file)[0]
    $endDate = (get-Date-Of-file $file)[1]
    
    #Write-Host "`:$BeginningTime`:$EndTime`:$BeginningDate`:$endDate"

    #if($HaveHeader-eq'no'){
    
    #    $Height = 
    #}

    $PrimaryXMLFile = @"
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>saeid</Author>
  <LastAuthor>saeid</LastAuthor>
  <Created>2023-09-13T16:21:21Z</Created>
  <LastSaved>2023-09-13T17:29:14Z</LastSaved>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>4695</WindowHeight>
  <WindowWidth>15345</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
  </Style>
  <Style ss:ID="s64">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders/>
   <Font ss:FontName="Vazir" x:Family="Swiss" ss:Color="#000000" ss:Bold="1"/>
   <Interior ss:Color="#ffffff" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="Vazir" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#ffffff" ss:Pattern="Solid"/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Vazir FD-WOL" x:Family="Swiss" ss:Size="11"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s67">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Vazir" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s68">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Wingdings 2" x:CharSet="2" x:Family="Roman" ss:Size="18"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
    <Style ss:ID="s418">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Vazir FD-WOL" x:Family="Swiss" ss:Size="11"
    ss:Color="#000000"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1" ss:RightToLeft="1">
  <Table ss:ExpandedColumnCount="4" ss:ExpandedRowCount="2" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s62" ss:DefaultRowHeight="15">
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="31.5"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="96.75"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="99"/>
   <Column ss:StyleID="s62" ss:Width="90"/>
   <Row ss:AutoFitHeight="0" ss:Height="37.5">
    <Cell ss:MergeAcross="3" ss:StyleID="s64"><Data ss:Type="String">لیست افرادی که در تاریخ $BeginningDate تا تاریخ $endDate حضور دارند بازه : $BeginningTime الی $EndTime</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="55.5">
    <Cell ss:StyleID="s65"><Data ss:Type="String">ردیف</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">نام</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">نام خانوادگی</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">کد پرسنلی</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <VerticalResolution>0</VerticalResolution>
    <NumberofCopies>0</NumberofCopies>
   </Print>
   <Selected/>
   <DisplayRightToLeft/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>5</ActiveRow>
     <ActiveCol>2</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
"@

    $PrimaryXMLFile|Out-File (outputXMLFile 1 $DateOutput)

}

$Header = (Get-heightWidth-andHeader -configFile 'config.conf')[2]

set-primartXMLFile -file $filesArray.values[0] -primaryXMLFile $PrimaryXMLFile -DateOutput $DateOutput -configFile 'config.conf' -HaveHeader $Header

#Set informations.xml file to XML file
$pe=0
forEach($index in $informations.Keys|sort){
  $name = ($informations.$index)[0]
  $family = ($informations.$index)[1]
  $number = ($informations.$index)[2]
  AddRow $(outputXMLFile 1 $DateOutput) $(outputXMLFile 1 $DateOutput) $index $name $family $number true

  $pe+=1
  $percentComplete = [math]::Round(($pe / $informations.keys.Count) * 100, 2)
  Show-ProgressBar -Caption "Set 'informations.xml' file to output file" -PercentComplete $percentComplete
}
Write-Host -b Black ""
Write-Host -b Black -f Green "process has been Completed!"

$Dates = Add_Dates 'config.conf'
#$Dates
$Rows = New-Object System.Collections.ArrayList
Write-Host -b Black ""

:fileLoop forEach ($file in $filesArray.keys){
    [xml]$xml=$filesArray.Item($file)
    $Rows0 = New-Object System.Collections.ArrayList
    $Rows00 = $xml.Workbook.Worksheet.Table.Row.forEach{if($_.Cell.Count -gt 0){$Rows0.Add($_)|Out-Null}}
    $pe=0

    forEach($Row0 in $Rows0){
    
      if([int]$Row0.Cell.Count -ne 0 -and
      $Row0.Cell[1].InnerText -as [long] -isnot [long] -and
      $Row0.Cell[2].InnerText -as [long] -isnot [long] -and
      $Row0.Cell[3].InnerText -as [long] -is    [long]){
      
         $Rows.Add($Row0)|Out-Null

      }else{
         
         switch($true){
            
            
             {$Row0.Cell[1].InnerText -as [long] -is [long]-and $Row0.Cell[1].InnerText -ne 'نام خانوادگی'}{
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    #[Console]::WriteLine([String]::Empty.PadLeft([Console]::WindowWidth))
                    Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    Write-Host -NoNewline -b Black -f Red "At Row:$($Rows0.IndexOf($Row0)+1) Cell:2 in the file `"$file`" the Text of the Cell is a number, while it's should be a String`n"
                    $progress = $false
            
             }{$Row0.Cell[2].InnerText -as [long] -is [long] -and $Row0.Cell[2].InnerText -ne 'نام'}{
                if($progress-eq $true){
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    Write-Host -NoNewline -b Black -f Red "At Row:$($Rows0.IndexOf($Row0)+1) Cell:3 in the file `"$file`" the Text of the Cell is a number, while it's should be a String`n"
                    $progress = $false

                }else{
                    Write-Host -NoNewline -b Black -f Red "At Row:$($Rows0.IndexOf($Row0)+1) Cell:3 in the file `"$file`" the Text of the Cell is a number, while it's should be a String`n"
                }
            
            }{$Row0.Cell[3].InnerText -as [long] -isnot [long] -and $Row0.Cell[3].InnerText -ne 'کد پرسنلي' -and $Row0.Cell[3].InnerText -ne 'کد پرسنلی'}{
                if($progress-eq $true){
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
                    [Console]::SetCursorPosition(0, [Console]::CursorTop)
                    Write-Host -NoNewline -b Black -f Red "At Row:$($Rows0.IndexOf($Row0)+1) Cell:4 in the file `"$file`" the Text of the Cell is a String while it's should be a Number`n"
                    $progress = $false

                }else{
                    Write-Host -NoNewline -b Black -f Red "At Row:$($Rows0.IndexOf($Row0)+1) Cell:4 in the file `"$file`" the Text of the Cell is a String while it's should be a Number`n"
                }
            
            }
         
         }
      
      }
          $pe+=1
          $percentComplete = [math]::Round(($pe / $Rows0.Count) * 100, 2)
          Show-ProgressBar -Caption "check and input `"$file`" file" -PercentComplete $percentComplete
          $progress=$true
    }
      Write-Host -b Black ""
      Write-Host -b Black -f Green "process has been Completed!"
}
$presentDate = @{}
$pe=0
$informations_sorted = $informations.Keys|sort

forEach ($Date in $Dates){
   $present=New-Object system.collections.generic.list[string]
   $p = 'F'
   $GregorianDate = Convert-ShamsiToGregorian $Date
   $date0 = ($Date-Split"/")[1]+'/'+($Date-Split"/")[2]

   forEach($index in $informations_sorted){
      $number = ($informations.$index)[2]
       
      :RowLoop forEach($Row in $Rows){
           $RowDate = $Row.Cell[0].InnerText

           #if([int]($date0-split'/')[0] -lt [int]($Row.cell[0].InnerText-split'/')[0]){
           #"...................Break......................1"
           #    $Row.cell[0].InnerText
           #    $date0
           #    [int]($date0-split'/')[0]
           #    [int]($Row.cell[0].InnerText-split'/')[0]
           #"...................Break......................2"
           #    Break RowLoop
           #}
           
           <#else#>if($Row.Cell[0].InnerText -eq $date0){
           
           #  "AAAAAAAAAAAAAAAA"
           #  $Row.cell[0].InnerText
           #  $date0
           #  [int]($date0-split'/')[0]
           #  [int]($Row.cell[0].InnerText-split'/')[0]
           #  "BBBBBBBBBBBBBBBBB"
              if($Row.Cell[3].InnerText -eq $number){

                try{
                
                   $present.Add($number)
                   
                   #Write-Host -f Green -b Black "the $number number added"
                   $p = 'T'
                   
                }catch{
                   
                   if($progress-eq $false){Write-Host -NoNewline -f Red -b Black "the `"$number`" number used for more than of one person"}
                   else{
                       [Console]::SetCursorPosition(0, [Console]::CursorTop)
                       Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
                       [Console]::SetCursorPosition(0, [Console]::CursorTop)
                       Write-Host -NoNewline -f Red -b Black "the number `"$number`" used for more than of one person`n"
                       $progress = $false
                   }
                   
                   break RowLoop
                }
              }
           }
        }
   }
          



   #AddCloumn "temp.xml" "temp.xml" $Date $present

   try{
       $presentDate.Add($Date,$present)
   }catch{
       if($progress-eq $false){Write-Host -NoNewline -f red -b Black "The date $Date is repeated"}
       else{
           [Console]::SetCursorPosition(0, [Console]::CursorTop)
           Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
           [Console]::SetCursorPosition(0, [Console]::CursorTop)
           Write-Host -NoNewline -f red -b Black "The date $Date is repeated`n"
           $progress = $false
       }
   }
   if($p -eq 'F'){
       if($progress-eq $false){Write-Host -NoNewline -f red -b Black "WARNING: Nobody was presented in Date: $Date"}
       else{
           [Console]::SetCursorPosition(0, [Console]::CursorTop)
           Write-Host -NoNewline -b Black ([String]::Empty.PadLeft([Console]::WindowWidth))
           [Console]::SetCursorPosition(0, [Console]::CursorTop)
           Write-Host -NoNewline -f red -b Black "WARNING: Nobody was presented in Date: $Date`n"
           $progress = $false
       }
   }
     #Write-Host -f green -b Black "the date: $Date added successfully"
     #sleep

   $pe+=1
   $percentComplete = [math]::Round(($pe / $Dates.Count) * 100, 2)
   Show-ProgressBar -Caption "set dates' presents" -PercentComplete $percentComplete
   $progress = $true
}

# Sort the hashtable by keys (dates)
$presentDate = $presentDate.GetEnumerator() | Sort-Object -Property Key | ForEach-Object {
    [ordered]@{ $_.Key = $_.Value }
}
#$presentDate = $sortedHashtable $presentDate.keys
Write-Host -b Black ""
Write-Host -b Black -f Green "process has been Completed!"
AddCloumn $(outputXMLFile 1 $DateOutput) $(outputXMLFile 1 $DateOutput) $presentDate

function set_Height($XMLFile,$out_XMLFile,[int]$HeightNumber){

   $data = [System.IO.File]::ReadAllText($XMLFile)
   [xml]$xml = $data
   $Rows = $data -split '</Row>'
   [int]$RowNumber = $xml.Workbook.Worksheet.table.ExpandedRowCount
   [int]$RowCount  = $xml.workbook.Worksheet.Table.ExpandedRowCount
   $newdata = $data

   #$pe=0
   for([int]$i=$HeightNumber+1; $i-lt $RowCount-1; $i+=$HeightNumber){
     $newdata = ($newdata -split $Rows[$i])[0]+
       $Rows[$i]+'</Row>'+ '<Row'+($Rows[0]-split'<Row')[1]+'</Row>'+$Rows[1]+
       ($newdata -split $Rows[$i])[1]

     $newdata=($newdata -split "ExpandedRowCount=`"")[0]+"ExpandedRowCount=`""+[String]($RowNumber=$RowNumber+2)+
     "`" x:FullColumns="+($newdata -split "`" x:FullColumns=")[1]

     #($data -split $splited[$i])[0] | Out-File $out_XMLFile

     $newdata | Out-File $out_XMLFile
      
     #sleep
     #$pe+=1
     $percentComplete = [math]::Round(($i / ($RowCount-1)) * 100, 2)
     Show-ProgressBar -Caption "Set Height" -PercentComplete $percentComplete
   }

   Write-Host -b Black ""
   Write-Host -b Black -f Green "process has been Completed!"
}

function set_Width($XMLFile,$out_XMLFile,[int]$widthNumber){

   $data = [System.IO.File]::ReadAllText($XMLFile)
   [xml]$xml = $data
   [int]$ColumnCount = $xml.workbook.Worksheet.Table.ExpandedColumnCount
   $Rows = $data -split '</Row>'
   $newdata=$data

   [xml]$xml = $newdata

   if(!(($ColumnCount-4) -lt $WidthNumber)){

    forEach($Row in $Rows){
      $newRow=""
      $cells = $Row -split '<Cell' -split '</Cell>'
      $addCell = '<Cell'+$cells[1]+"</Cell>`n"+'<Cell'+$cells[3]+"</Cell>`n"+'<Cell'+$cells[5]+"</Cell>`n"+'<Cell'+$cells[7]+"</Cell>`n"
      [int]$newColumnCount = $ColumnCount
      $o1 = 0
      if(1-eq1){
         for([int]$i=2*($widthNumber+4); $i-lt 2*$ColumnCount; $i+=(2*$widthNumber)){
           
           #$newRow = ($Row -split $cells[$i])[0]+$cells[$i]+$add+($Row -split $cells[$i])[1]
           #if ($o1%2 -eq 0){

              try{
               $cells[$i] = $cells[$i]+$addCell
               $newColumnCount+=4
              }catch{
              #$Row
              
              }

           #}
           $o1++
         }
      }
      $o=0
      forEach($cell in $cells){

        if ($o%2 -ne 0){
            $newRow+='<Cell'+$cell+'</Cell>'
        }else{
            $newRow+=$cell
        }
        $o++
      }

      if($newRow.contains("MergeAcross=`"")){

            $newRow = ($newRow -split "MergeAcross=`"")[0]+"MergeAcross=`""+[string]($widthNumber+3)+
            "`" ss:StyleID=`""+((($newRow -split "MergeAcross=`"")[1]) -split "`" ss:StyleID=`"")[1]
            
            $cellM = (($newRow -split '<Cell')[1] -split '</Cell>')[0]

         for([int]$i=$widthNumber+4; $i-lt $ColumnCount; $i+=$widthNumber){
         #$ColumnCount
            if(($ColumnCount-$i)-lt $widthNumber){


                $cellM2='<Cell'+(($newRow -split '<Cell')[1]+$newRow -split '</Cell>')[0]

                $newRow +="    "+($cellM2 -split "MergeAcross=`"")[0]+"MergeAcross=`""+[string]($ColumnCount-$i+3)+
                "`" ss:StyleID=`""+($cellM2 -split "`" ss:StyleID=`"")[1]+"</Cell>`n"
            
            }else{
                $newRow += " <Cell"+$cellM+"</Cell>`n"
               
            }
            

         }
      }

      #"AAAAAAAAAAAAAAAAAAAAAAAAAAAAaa"
      #$newRow
      #"BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB"
      $newdata = $newdata.replace($Row,$newRow)
      $pe+=1
      $percentComplete = [math]::Round(($pe / $Rows.Count) * 100, 2)
      Show-ProgressBar -Caption "Set width" -PercentComplete $percentComplete

   }
   Write-Host -b Black ""
   Write-Host -b Black -f Green "process has been Completed!"

   $Columns = $newdata -split '<Column'

   $pe=0
   for([int]$i=$widthNumber+4; $i-lt $ColumnCount; $i+=$widthNumber){
   
       $Columns[$i] = $Columns[$i]+'<Column'+$Columns[1]+'<Column'+$Columns[2]+'<Column'+$Columns[3]+'<Column'+$Columns[4]
      
       $columnNumber = (($Columns[0] -split "ExpandedColumnCount=`"")[1] -split "`" ss:ExpandedRowCount")[0]

       $Columns[0] = ($Columns[0] -split "ExpandedColumnCount=`"")[0]+"ExpandedColumnCount=`""+
     +([int]$columnNumber+4)+"`" ss:ExpandedRowCount"+($Columns[0] -split "`" ss:ExpandedRowCount")[1]

        $pe+=1
        $percentComplete = [math]::Round(($pe / $ColumnCount) * 100, 2)
        Show-ProgressBar -Caption "Set Columns" -PercentComplete $percentComplete
   }
   Write-Host -b Black ""
   Write-Host -b Black -f Green "process has been Completed!"

   $pe=0
   $newdata=''
   $o3=0
   forEach($Column in $Columns){
      

      if ($o3-eq 0){
        $newdata+=$Column

      }else{
        $newdata+='<Column'+$Column

      }
      $o3++

      $pe+=1
      $percentComplete = [math]::Round(($pe / $Columns.Count) * 100, 2)
      Show-ProgressBar -Caption "Set Columns Header" -PercentComplete $percentComplete
   }
  }else{
        Show-ProgressBar -Caption "Set width" -PercentComplete 100
        Write-Host -b Black ""
        Write-Host -b Black -f Green "process has been Completed!"

        Show-ProgressBar -Caption "Set Columns" -PercentComplete 100
        Write-Host -b Black ""
        Write-Host -b Black -f Green "process has been Completed!"

        Show-ProgressBar -Caption "Set Columns Header" -PercentComplete 100
  }
   Write-Host -b Black ""
   Write-Host -b Black -f Green "process has been Completed!"

   $newdata | Out-File $out_XMLFile


}

function set_width3($XMLFile,$out_XMLFile,[int]$widthNumber,[int]$HeightNumber){
   $Rows = New-Object System.Collections.ArrayList
   #$newdata = New-Object System.Collections.ArrayList
   #$RowHeader= New-Object System.Collections.ArrayList
   $data = [System.IO.File]::ReadAllText($XMLFile)
   $Columns0 = $data -split '<Column'
   $xml = [xml]$data
   [int]$ColumnCount = $xml.workbook.Worksheet.Table.ExpandedColumnCount

 if(!(($ColumnCount-4) -lt $WidthNumber)){

   for($i=1; $i-le $widthNumber+4; $i++){
       $Columns+='<Column'+$Columns0[$i]
   }
   $data=$Columns0[0]+$Columns+($Columns0[-1]-split ($Columns0[-1]-split'/>')[0]+'/>')[1]

   $LRow0 = [int]$xml.workbook.Worksheet.Table.Row[0].cell.Count
   $LRow1 = [int]$xml.workbook.Worksheet.Table.Row[1].cell.Count
   $Rows.addRange([string[]]($data-split '</Row>'))
   $Header = ($Rows[0] -split"<Row ")[0]
   $Footer = $Rows[-1]
   $Rows[0]="<Row "+($Rows[0]-split"<Row ")[1]
   $Rows.removeAt($Rows.Count-1)
   $Header=($Header -split "ExpandedRowCount=`"")[0]+"ExpandedRowCount=`""+[String]($Rows.Count*$LRow0)+
     "`" x:FullColumns="+($Header -split "`" x:FullColumns=")[1]
   $newdata+=$Header

   $pe=0
   for($t=0;$t-lt $LRow0;$t++){
       $Merge=[int]$xml.Workbook.Worksheet.Table.Row[0].cell[$t].MergeAcross+1
       $newdata0=''
         
       forEach($Row in $Rows){
          $Cells = New-Object System.Collections.ArrayList
          $Cells0 = $Row-split '<Cell ss'
          $Cells.AddRange($Cells0[1..($Cells0.Count-1)].ForEach{'<Cell ss'+$_})
          #$Cells.Add($Cells0[-1])|Out-Null
          #$Cells.Add($Cells0[0])|Out-Null
          #$RowHeader.Add($Cells[0])|Out-Null
          $RowHeader = $Cells0[0]
          #$Cells.RemoveAt(0)
          #$Cells.Count
          #$RowHeader
          if($Cells.Count-eq$LRow1){
             $newdata0+=$RowHeader+$Cells[$Merge0..($Merge0+$Merge-1)]+'</Row>'
          }elseif($Cells.Count-eq$LRow0){
             $newdata0+=$RowHeader+$Cells[$t]+"`n</Row>"
          }else{
             #Write-Error "set_width3$Cells"
          
          }
    
       
       }
       [int]$Merge0+=[int]$Merge
       $newdata+=$newdata0

       $pe+=1
       $percentComplete = [math]::Round(($pe / $LRow0) * 100, 2)
       Show-ProgressBar -Caption "Add Rows" -PercentComplete $percentComplete
    }
 }else{
    $newdata = $data
    Show-ProgressBar -Caption "Add Rows" -PercentComplete 100
    
 }
    Write-Host -b Black ""
    Write-Host -b Black -f Green "process has been Completed!"

    $newdata+=$Footer
    $newdata | Out-File $out_XMLFile
}

$HeightWidth=Get-heightWidth-andHeader 'config.conf'
#"Width:$($HeightWidth[1])"
#"Height:$($HeightWidth[0])"
set_Height -XMLFile $(outputXMLFile 1 $DateOutput) -out_XMLFile $(outputXMLFile 2 $DateOutput) -HeightNumber $HeightWidth[0]
set_Width -XMLFile $(outputXMLFile 2 $DateOutput) -out_XMLFile $(outputXMLFile 2 $DateOutput) -widthNumber $HeightWidth[1]
#set_Width2 -XMLFile "temp2.xml" -out_XMLFile "temp3.xml" -widthNumber '1' -HeightNumber '50'
set_width3 -XMLFile $(outputXMLFile 2 $DateOutput) -out_XMLFile $(outputXMLFile 3 $DateOutput) -widthNumber $HeightWidth[1] -HeightNumber $HeightWidth[0]
ExitProgram