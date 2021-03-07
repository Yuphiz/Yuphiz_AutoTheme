<#
.声明
脚本：壁纸高级版_Yuphiz 
版本：v0.1
作者：Yuphiz
本脚本可以帮助 Windows10 自动切换深浅色主题的壁纸
凡用此脚本从事法律不允许的事情的，均与本作者无关
此脚本许可采用 GPL-3.0-later 协议
#>

param (

)

function get-ScheduledTask-State {
      param (
            $TaskName,
            $RootPath = "\"
      )                  
      foreach ($OneOf in $TaskName) {
            $service = new-object -com("Schedule.Service")
            $service.connect()
            $rootFolder = $service.Getfolder($RootPath)
            $taskDefinition = $service.NewTask(0)
            if ($(try{$rootFolder.gettask($OneOf).enabled} catch{$false})){
                  return $True
            }
      }
      return $False   
}


$OthersTaskArray= "实时壁纸\登录",`
                  "实时壁纸\唤醒", `
                  "实时壁纸\日出前后", 
                  "实时壁纸\日落前后", `
                  "实时壁纸\日间", `
                  "实时壁纸\夜间", `
                  "实时壁纸\多组文件夹支持", `
                  "Windows聚焦壁纸\Windows聚焦壁纸", `
                  "每日bing壁纸\每日bing壁纸", `
                  "每日bing壁纸\登录", `
                  "每日bing壁纸\唤醒"
$IsHaveOtherWallpaperScript = get-ScheduledTask-State $OthersTaskArray "\YuphizScript\$env:username"

if ($IsHaveOtherWallpaperScript) { exit } #如果启用其他壁纸模块，则直接退出不执行下面的更换壁纸

#脚本所在路径，如果为空则选择当前工作路径
$PathScriptWork = $PSScriptRoot;if ($PathScriptWork -eq "") {$PathScriptWork=(get-location).path}
$title = "自动主题配色"
$popup=new-object -comobject wscript.shell

$FileConfig = "$PathScriptWork\壁纸高级版_Yuphiz.Json"
if (!(Test-Path $FileConfig)){
      @'
{
      "备注":"地址路径的单斜杠(\)要写成双斜杠(\\)"
,     "替换图片方式":"多壁纸组"
,     "单壁纸组选项":{
             "浅色主题壁纸":""
      ,      "深色主题壁纸":""
      }
,     "多壁纸组和轮播选项":{
             "壁纸包路径":""
      ,      "多壁纸组选项":{
                  "随机壁纸组":"开"
            }
      ,      "轮播选项":{
                  "轮播时间_秒":7200
            }
      }
,     "默认壁纸契合度":"拉伸"
,     "自定义壁纸契合度":{
            "untitle":"填充"
      }
,     "支持图片类型":[".jpg",".jpeg",".png",".bmp"]
}
'@ | set-content $FileConfig
}
try {
      $Config=get-content $FileConfig | ConvertFrom-Json
}catch{
      $null = $popup.popup("      $FileConfig `n`r
$($error[0])",0,"配置文件出错",16);exit
}

# 替换图片方式，0为单张图片，只在日出日落替换白天/夜晚图片，1为壁纸组，每天替换不同组的图片，2为随机轮播，白天轮播白天的图片，夜晚轮播夜晚的图片。设置只在下次运行时生效，也可以右键powershell运行立即生效
$WallpaperRunBy = $Config.替换图片方式

# 替换图片方式设为 0 时生效的设置，只会在日出日落时换日出日落图片
#单张浅色主题壁纸
$DayWallpaper = $Config.单壁纸组选项.浅色主题壁纸
#单张深色主题壁纸
$NightWallpaper = $Config.单壁纸组选项.深色主题壁纸


# 替换图片方式设为 1 或者 2 时生效的设置，除了在日出日落时换日出日落图片，还有:
# 设置为 1 时 每天从一组日出日落图片中自动选择，即在 0 的基础上，每天换一组图片，需要文件名支持
# 设置为 2 时 白天轮播白天的图片，夜晚轮播夜晚的图片，需要文件名支持
# 日出日落图片组文件夹，空值则为根目录
$WallpaperFolder = $Config.多壁纸组和轮播选项.壁纸包路径

# 顺时轮播文件组，0为顺序，1为随机
$TurnOnRandomFolder = $Config.多壁纸组和轮播选项.多壁纸组选项.随机壁纸组

# 替换图片方式设为 2 时生效的设置
#轮播周期，单位是秒，不能少于 60 s
$CarouselTime = $Config.多壁纸组和轮播选项.轮播选项.轮播时间_秒

#支持图片类型
$FormatSupported = $Config.支持图片类型

#默认壁纸契合度
$defaultWallpaperStyle = $Config.默认壁纸契合度

#特殊壁纸契合度，以下面命名的文件夹下的壁纸（不包括_1、_2）都会以自定义的方式设为壁纸
$WallpaperStyle = $Config.自定义壁纸契合度




$PathExtensionsLeaf = Split-Path $PathScriptWork
$VbsLauncher = "$PathExtensionsLeaf\自动主题配色_辅助启动.vbs"
while ((Test-Path($VbsLauncher)) -eq $False) {
      $PathExtensionsLeaf = Split-Path $PathExtensionsLeaf
      $VbsLauncher = "$PathExtensionsLeaf\自动主题配色_辅助启动.vbs"
}

switch ($WallpaperRunBy) {
      0 { 
            if ($DayWallpaper -eq ""){
                  $DayWallpaper = "$PathScriptWork\壁纸包\surface\Surface_Laptop_3_03_1.jpg"
            }
            if ($NightWallpaper -eq ""){
                  $NightWallpaper = "$PathScriptWork\壁纸包\surface\Surface-Wallpaper-4500x3000-67826_2.jpg"
            }
       }
      {$_ -eq 1 -or $_ -eq 2} {
            if ($carouselTime -le 60 ){ $carouselTime = 60 }
            if ($WallpaperFolder -eq "" ) {
                  $WallpaperFolder = "$PathScriptWork\壁纸包"
                  if (!(Test-Path $WallpaperFolder)){
                        ni $WallpaperFolder -ItemType Directory -Force
                  }
            }else{
                  if (!(Test-Path $WallpaperFolder)){
                       $null = $popup.Popup("   找不到路径 ,请重新设置 `n`r
    $WallpaperFolder ",0,"警告",16 + 4096)
                        exit
                  }
            }
      }
      default {
            $null = $popup.popup("   替换图片方式的值只能是0或1或2 `n`r
    0：单壁纸组，只在日出日落替换白天/夜晚图片，`n`r
    1：多壁纸组，每天替换不同组的图片，`n`r
    2：随机轮播，白天轮播白天的图片，夜晚轮播夜晚的图片
    ",0,"出错",16 + 4096)
            exit
      }
}



$SystemTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").SystemUsesLightTheme
$AppTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme
$LastWallpaper = (Get-ItemProperty -path "registry::HKEY_CURRENT_USER\Control Panel\Desktop").wallpaper

      
function WirteOrReadDataFromSchtask {
      param (
            $WirteOrRead,
            $Data = $null,
            $RepetitionInterval = $null
      )
      
      $service = new-object -com("Schedule.Service")
      $service.connect()
      $rootFolder = $service.Getfolder("\YuphizScript")
                  
      $taskDefinition = $service.NewTask(0)
                  
      if ($WirteOrRead -eq "Write"){
            $Settings = $taskDefinition.Settings
            $Settings.StartWhenAvailable = $True
            $Settings.DisallowStartIfOnBatteries = $false
            $Settings.ExecutionTimeLimit= "PT5M"
            
            $triggers = $taskDefinition.Triggers
            $TriggerTypeLogin=9
            $trigger = $triggers.Create($TriggerTypeLogin)
            $trigger.UserId = $env:username
            $trigger.delay = "PT90S"
            $trigger.Repetition.Interval = $RepetitionInterval
            $trigger.Repetition.Duration = $null
            
            $InfoTask = $taskDefinition.RegistrationInfo
            $InfoTask.Description = $Data
            
            $Action = $taskDefinition.Actions.Create(0)
            $Action.Path = "wscript"
            $Action.Arguments= `
            "`"$VbsLauncher`" --Wallpaper"
            
            $CreateOrUpdateTask = 6
            $null=$rootFolder.RegisterTaskDefinition( `
            "$env:username\$title\壁纸轮播",`
            $taskDefinition,$CreateOrUpdateTask,$null,$null, 3)
            
      }elseif($WirteOrRead -eq "Read"){
            if (schtasks /query /tn YuphizScript\$env:username\$title\壁纸轮播 2>$null){
                  $XmlTasks= `
                  [xml]($rootFolder.gettask("$env:username\$title\壁纸轮播").xml)
                  $DescriptionTasks=$XmlTasks.Task.RegistrationInfo.Description
                  return $DescriptionTasks
            }
      }
      
}



function Get-Folder {
      param (
            $PathOfFolder
      )
      $FilterFolders = (ls $PathOfFolder -Directory -Depth 1) | ?{$_.Name.IndexOf("__") -ne 0}
      $AllFolders = @()
      $AllFolders += $PathOfFolder
      foreach ($Oneof in $FilterFolders) {
            $path = ($Oneof.FullName.split("\"))[-1,-2,-3] | ?{ $_.indexof("__") -eq 0}
            if ($path.count -eq 0) {
                  $AllFolders += $Oneof
            }
      }
      
      $Folders=@()
      foreach ($Oneof in $AllFolders) {
            $Filter=(ls $Oneof.fullname -File | ?{$FormatSupported -contains $_.Extension})
            if ($Filter.count -ge 1){
            $i=$ii=0
            foreach ($oneof2 in $Filter.basename) {
                  if ($oneof2.indexof("_1") -eq "$OneOf2".length-2 -and "$OneOf2".length -ge 2){
                        $i++
                        if ($i -gt 1){ break }
                  }elseif ($oneof2.indexof("_2") -eq "$OneOf2".length-2 -and "$OneOf2".length -ge 2) {
                        $ii++
                        if ($i -gt 1){ break }
                  }
            }
            $FolderName = Split-Path ($Oneof.fullname) -leaf
            if ( $i -eq 1 -and $ii -eq 1 -and $FolderName.IndexOf("__") -ne 0) {
                  $Folders += $Oneof.fullname
            }
            }
      }
return $Folders
}


Function Get-NewWallpaperFolder {
      $OldWallpaperFolder = Split-Path -parent $LastWallpaper
      # $IsHavaLastWallpaper = (Test-Path $OldWallpaperFolder)

      [object[]]$Folders = Get-Folder $WallpaperFolder
      if ($Folders.Count -eq 0) {
            $null = $popup.Popup("   找不到符合的图片文件 `n`r
    $WallpaperFolder `n`r
    即将退出壁纸高级版",0,"警告",16 + 4096)
            exit
      }
      
      $DateOld = (WirteOrReadDataFromSchtask "Read").split(",")[0]
      $Today = (Get-Date -Format "yyyy-MM-dd")
      try {$null=(get-date $DateOld)}catch{if ($error[0] -match "Date"){$DateOld = 0}}
      if ($DateOld -eq $null -or (New-TimeSpan $DateOld $Today).Days -ge 1) {
            WirteOrReadDataFromSchtask "Write" "$Today,$CarouselTime"

            for ($i=0;$i -lt $Folders.count;$i++) { 
                  if ($OldWallpaperFolder -eq $Folders[$i]){
                        break
                  }
            }
            if ($TurnOnRandomFolder -eq "开" -and $Folders.count -gt 1) {
                  do {
                        $newI=random(0..$($Folders.count-1))
                  } until ( $newI -ne $i )
            }elseif ($TurnOnRandomFolder -eq "关" -or $Folders.count -le 1) {
                  if ($newI++ -gt $Folders.count) {
                        $newI = 0
                  }else {
                        $newI++
                  }
            }
            
            $newWallpaper = $Folders[$newI]
            
      }else{
            foreach ($oneof in $Folders){
                  if ($oneof -eq $OldWallpaperFolder) {
                        $newWallpaper = $OldWallpaperFolder
                        break
                  }
            }
            if ($newWallpaper -eq $null){
                  $newWallpaper = $Folders[0]
            }
      }
Return $newWallpaper
}



function Get-NewWallpaper {
      param (
            $FolderPath,
            $Theme
      )
      $NewWallpaper = switch ($Theme) {
            1 {
                  ((ls $FolderPath -file) | ?{$_.BaseName.IndexOf("_1") -eq $_.BaseName.Length-2 -and $_.basename.length -ge 2}).FullName 
            }
            0 { 
                  ((ls $FolderPath -file) | ?{$_.BaseName.IndexOf("_2") -eq $_.BaseName.Length-2 -and $_.basename.length -ge 2}).FullName
            }
      }
return $NewWallpaper
}



function Get-NewWallpaperFromAll {
      param (
            $FolderPath,
            $Theme,
            $allwallpapers=@(),
            $FormatSupported =@( ".jpg",".jpeg",".png",".bmp")
      )
      [object[]]$AllPictureFile = (ls $FolderPath -Depth 2) | ?{$FormatSupported -contains $_.Extension }

      foreach ($Oneof in $AllPictureFile) {
            $path = ($Oneof.FullName.split("\"))[-1,-2,-3] | ?{ $_.indexof("__") -eq 0}
            if ($path.count -eq 0) {
                  $allwallpapers += $Oneof
            }
      }

      if ($Theme -eq 1) {
            [object[]]$DayWallpapers = $allwallpapers | ? {$_.basename.indexof("_1") -eq $_.basename.length-2 -and $_.basename.length -ge 2}
            
            switch ($DayWallpapers.count) {
                  {$_ -gt 1} { 
                        do {
                        $Number = Random(0..$($DayWallpapers.Count-1))
                        } until ($DayWallpapers.fullname[$Number] -ne $LastWallpaper)
                        $NewWallpaperFromAll = $DayWallpapers.fullname[$Number] 
                  }
                  1 {
                        $NewWallpaperFromAll = $DayWallpapers[0].fullname
                  }
                  {$_ -lt 1} {
                        $null = $popup.Popup("   找不到符合的图片文件 `n`r
    $WallpaperFolder `n`r
    即将退出壁纸高级版",0,"警告",16 + 4096)
                  exit
                  }
            }

      }elseif ($Theme -eq 0) {
            [object[]]$NightWallpapers = $allwallpapers | ? {$_.basename.indexof("_2") -eq $_.basename.length-2 -and $_.basename.length -ge 2}

            switch ($NightWallpapers.count) {
                  {$_ -gt 1} { 
                        do {
                              $Number = Random(0..$($NightWallpapers.Count-1))
                        } until ($NightWallpapers.fullname[$Number] -ne $LastWallpaper)
                        $NewWallpaperFromAll = $NightWallpapers.fullname[$Number]
                   }
                  1 {
                        $NewWallpaperFromAll = $NightWallpapers[0].fullname
                   }
                  {$_ -lt 1} {
                        $null = $popup.Popup("   找不到符合的图片文件 `n`r
    $WallpaperFolder `n`r
    即将退出壁纸高级版",0,"警告",16 + 4096)
                  exit
                   }
            }
      }
return $NewWallpaperFromAll
}



switch ($WallpaperRunBy) {
      0 {
            $Wallpaper = switch ($AppTheme) {
                  1 { $DayWallpaper }
                  0 { $NightWallpaper }
            }
            $null = schtasks /change /disable /tn YuphizScript\$env:username\$title\壁纸轮播
      }
      1 {
            $NewWallpaperFolder = Get-NewWallpaperFolder
            $Wallpaper = Get-NewWallpaper $NewWallpaperFolder $AppTheme
            $null = schtasks /change /disable /tn YuphizScript\$env:username\$title\壁纸轮播
      }
      2 {
            $CarouselTimeData =  (WirteOrReadDataFromSchtask "Read").split(",")[1]
            $IsEnabled = get-ScheduledTask-State "\YuphizScript\$env:username\$title\壁纸轮播"
            if ($CarouselTimeData -eq $null -or $CarouselTimeData -ne $CarouselTime -or $IsEnabled -eq $False) {
                  $Today = (Get-Date -Format "yyyy-MM-dd")
                  WirteOrReadDataFromSchtask "Write" "$Today,$CarouselTime" "PT$($CarouselTime)S"
            }
            $Wallpaper = Get-NewWallpaperFromAll $WallpaperFolder $AppTheme
      }
}



Function UpdateWallpaper {
      param (
            $Wallpaper,
            $Style = $defaultWallpaperStyle
      )
      if ($LastWallpaper -ne $Wallpaper) {
            If (test-path $Wallpaper) {
Add-Type @"
using System;
using System.Runtime.InteropServices;
namespace Wallpaper
{
      public class Setter {
            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
            public static void SetWallpaper ( string path) {
                  SystemParametersInfo( 20, 0, path, 1 | 2 );
            }
      }
}
"@
                  switch ($Style) {
                        "填充" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 0 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 10 /f;
                        }
                        "适应" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 0 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 6 /f;
                        }
                        "拉伸" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 0 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 2 /f;
                        }
                        "平铺" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 1 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 0 /f;
                        }
                        "居中" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 0 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 0 /f;
                        }
                        "跨区" {
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v TileWallpaper /t REG_SZ /d 0 /f;
                              $null = reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d 22 /f;
                        }
                  }
                  [Wallpaper.Setter]::SetWallpaper($Wallpaper)
            }else{
                  $null=$popup.Popup("壁纸更新失败，请检查文件是否存在",0,$null,4096)
            }
      }else{
            "图片相同，不需要换"
      }
}

$WallpaperFit = $Wallpaper
for ($i=0;$i -lt 3; $i++){
      $WallpaperFolderName = Split-Path $WallpaperFit -leaf
      $WallpaperFit = Split-Path $WallpaperFit
      if ($WallpaperStyle[$WallpaperFolderName] -ne $null){
            $defaultWallpaperStyle = $WallpaperStyle[$WallpaperFolderName]
            break
      }
}

UpdateWallpaper $Wallpaper $defaultWallpaperStyle
# $null=$popup.Popup("更换壁纸",1,$null,4096)