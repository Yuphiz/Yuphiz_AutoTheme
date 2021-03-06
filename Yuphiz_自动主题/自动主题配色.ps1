<#
.声明
      脚本名称：Yuphiz_自动主题
      版本号：v1.1.1
      作者：Yuphiz
      本脚本可以帮助 Windows10 自动切换深浅色主题
      凡用此脚本从事法律不允许的事情的，均与本作者无关
      此脚本许可采用 GPL-3.0-later 协议
#>


param (
      [Parameter(Mandatory=$false)]$RunWith
)

# $time1 = get-date
# 环境配置
#脚本所在路径，如果为空则选择当前工作路径
$PathScriptWork = $PSScriptRoot;if ($PathScriptWork -eq "") {$PathScriptWork=(get-location).path}
$title="自动主题配色"

$popup=new-object -comobject wscript.shell


function set-MulTrigger-TaskService {
      param (
            [Parameter(Mandatory=$true)]$TaskName,
            [Parameter(Mandatory=$true)]$TriggersArray,
            [Parameter(Mandatory=$true)]$ActionPath,
            [Parameter(Mandatory=$true)]$ActionArguments,
            [Parameter(Mandatory=$true)]$Description,
            [Parameter(Mandatory=$true)]$WhatToDoTask,
            $RootPath = "\"
      )
      $service = new-object -com("Schedule.Service")
      $service.connect()
      $rootFolder = $service.Getfolder($RootPath)

      $taskDefinition = $service.NewTask(0)

      $Settings = $taskDefinition.Settings
      $Settings.StartWhenAvailable = $True
      $Settings.DisallowStartIfOnBatteries = $false
      $Settings.ExecutionTimeLimit= "PT5M"

      $triggers = $taskDefinition.Triggers

      $TriggerHashTable = @{}
      For ($i=0; $i -lt $TriggersArray.count;$i++) {
            switch ($TriggersArray[$i].Type) {
                  "Logon" {
                        $TypeLogon = 9
                        ($TriggerHashTable.$i) = $triggers.Create($typeLogon)
                        ($TriggerHashTable.$i).UserId = $env:username
                        ($TriggerHashTable.$i).Enabled = $TriggersArray[$i].Enable
                        ($TriggerHashTable.$i).delay = $TriggersArray[$i].delay
                        ($TriggerHashTable.$i).Repetition.Interval = $TriggersArray[$i].Interval
                        ($TriggerHashTable.$i).Repetition.Duration = $TriggersArray[$i].Duration
                        break
                  }
                  "Daily" {
                        $TypeDaily = 2
                        ($TriggerHashTable.$i) = $triggers.Create($TypeDaily)
                        ($TriggerHashTable.$i).StartBoundary = $TriggersArray[$i].StartTime
                        ($TriggerHashTable.$i).DaysInterval = $TriggersArray[$i].DaysInterval
                        ($TriggerHashTable.$i).Repetition.Interval = $TriggersArray[$i].Interval
                        ($TriggerHashTable.$i).Repetition.Duration = $TriggersArray[$i].Duration
                        ($TriggerHashTable.$i).Enabled = $TriggersArray[$i].Enable
                        break
                    }
                  "Event" { 
                        $TypeEvent = 9
                        ($TriggerHashTable.$i) = $triggers.Create($TypeEvent)
                        ($TriggerHashTable.$i).Subscription = $TriggersArray[$i].XML
                        ($TriggerHashTable.$i).Enabled = $TriggersArray[$i].Enable
                        break
                   }
            }
      }

      $InfoTask = $taskDefinition.RegistrationInfo
      $InfoTask.Description = $Description

      $Actions = $taskDefinition.Actions
      $Action = $Actions.Create(0)
      $Action.Path = $ActionPath
      $Action.Arguments= $ActionArguments
            
      $WhatToDo = switch ($WhatToDoTask) {
            "Update" { 4 }
            "CreateOrUpdate" { 6 }
      }
      $null=$rootFolder.RegisterTaskDefinition( $TaskName,$taskDefinition,$WhatToDo,$null,$null, 3)
}

function get-TaskService {
      Param (
            [Parameter(Mandatory=$true)]$TaskName,
            $RootPath = "\"
      )
      $Results = @()
      $service = new-object -com("Schedule.Service")
      $service.connect()
      $rootFolder = $service.Getfolder($RootPath)
      
      $taskDefinition = $service.NewTask(0)

      Foreach ($Oneof in $TaskName) {
            try{
                  $Result = $rootFolder.gettask($Oneof)
            }catch{
                  $Result = "任务计划不存在"
            }
            $Results += @{
                  (split-path $Oneof -leaf) = $Result
            }
      }
return $Results
}



function get-ConfigFromJson {
      $FileConfig="$PathScriptWork\$($title)_$($env:userdomain)_$($env:username).json"

      if (! (Test-Path $FileConfig)) {
            @"
{
"定位选项":{
      "网络自动定位":"开"
,     "手动定位经度":0
,     "手动定位纬度":0
      
,     "日出时间偏移_时":0
,     "日落时间偏移_时":0
       
}

,"主题颜色选项":{
      "切换系统颜色":"开"
,     "切换应用颜色":"开"
}

,"扩展选项":{
      "启用扩展":"开"
,     "自动关闭":"关,此功能暂时不可选"
,     "扩展参数":{
            "测 试":"bbb ccc"
      }
,     "异步处理最大线程数":"8,此功能暂时不可选"
,     "延迟扩展":[
            "UWP启动页颜色"
      ]
,     "延迟时间_毫秒":4000
,     "延迟扩展异步处理":"关,此功能暂时不可选"
}
,"其他选项":{
      "任务计划偏移时间_时":0.05
}
}
"@ | set-content $FileConfig
}

      try {
            $Config=get-content $FileConfig | ConvertFrom-Json
      }catch{
            $null = $popup.popup("      $FileConfig `n`r
      $($error[0])",0,"配置文件出错",16);exit
      }
return $Config
}

if ($RunWith -eq "DefaultCongif") {exit}
# 读取配置
$Config=get-ConfigFromJson



# 获取扩展
function Get-Extensions {
      param (
            $PathOfFolder,
            $Extensions=@(),
            $FormatSupported =@( ".ps1")
      )
      $Filter=(ls $PathOfFolder -File -Depth 2 | ?{$FormatSupported -contains $_.Extension}) 
            foreach ($Oneof in $Filter) {
                  $path = ($Oneof.FullName.split("\"))[-1,-2,-3] | ?{ $_.indexof("__") -eq 0}
                  if ($path.count -eq 0) {
                        $Extensions += $Oneof
                  }
            }

return $Extensions
}


# 运行延迟的扩展
function Run_Delay_Extension {
      param (
            $Delay_extensions
      )
      # $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, 2)
      # $RunspacePool.Open()

      foreach ($oneof in $Delay_extensions){
            Invoke-Command -ScriptBlock $oneof
      }
      # Read-Host
}



# 运行扩展
function RunExtension {
      # $time1=get-date

      $extensions=Get-Extensions "$PathScriptWork\扩展"

      if ($extensions.count -gt 0){
      $Delay_extensionsWithArgument = @()
      $Delay_extensionsWithoutArgument = @()
      
      $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, 8)
      $RunspacePool.Open()
      $jobObject=@()
      for ($i=0;$i -lt $extensions.count;$i++) {
            $name=$extensions[$i].BaseName
            Start-Sleep -Milliseconds 200
            if ($config.扩展选项.扩展参数.$name -ne $null) {
                  if ($config.扩展选项.延迟扩展 -contains $name) {
                        $Argument=$config.扩展选项.扩展参数.$name
                        $Delay_extensionsWithArgument += [scriptblock]::Create("powershell -ExecutionPolicy Bypass -File '$($extensions[$i].fullname)' $Argument")
                  }else{
                        $Argument=$config.扩展选项.扩展参数.$name
                        $file = [scriptblock]::Create("powershell -ExecutionPolicy Bypass -File '$($extensions[$i].fullname)' $Argument")
                        $PowerShell =[powershell]::Create()
                        $PowerShell.Runspacepool = $Runspacepool
                        [void]$PowerShell.AddScript($file)
                        $jobObject += $PowerShell.BeginInvoke()
                  }
                  
            }else {
                  if ($config.扩展选项.延迟扩展 -contains $name) {
                        $Delay_extensionsWithoutArgument += [scriptblock]::Create("powershell -ExecutionPolicy Bypass -File '$($extensions[$i].fullname)' ")
                  }else{
                        $file = [scriptblock]::Create("powershell -ExecutionPolicy Bypass -File '$($extensions[$i].fullname)' ")
                        $PowerShell =[powershell]::Create()
                        $PowerShell.Runspacepool = $Runspacepool
                        [void]$PowerShell.AddScript($file)
                        $jobObject += $PowerShell.BeginInvoke()
                  }
            }
      }
      # ($(get-date)-$time1).TotalSeconds
      # $time1=get-date
      
      foreach ($Oneof in $jobObject) {
            # ($jobObject | ?{$_.Result.IsCompleted -ne $true}).count
            $null=$Oneof.AsyncWaitHandle.WaitOne()
      }
      $PowerShell.RunspacePool.close()
      $PowerShell.Dispose()
      $RunspacePool.close()
      $RunspacePool.Dispose()
      #     write-host "Handles: $($(Get-Process -Id $PID).HandleCount) Memory: $($(Get-Process -Id $PID).PrivateMemorySize64 / 1mb) mb"
      [System.GC]::Collect()
      # ($(get-date)-$time1).TotalSeconds
      $DelayTime = $config.扩展选项.延迟时间_毫秒
      if ($DelayTime -lt 2000){
            $DelayTime = 5000
      }
      Start-Sleep -Milliseconds $DelayTime
      $Delay_extensionsWithArgument += $Delay_extensionsWithoutArgument
      Run_Delay_Extension $Delay_extensionsWithArgument
}
      # read-host
}


#后台版任务计划进程id写入更新、进程id读取
#═══════════════ 更新计划任务(后台数据) 开始 ═════════════════
function TaskService {
      param (
            $UpdateOrRead
      )
      $service = new-object -com("Schedule.Service")
      $service.connect()
      $rootFolder = $service.Getfolder("\YuphizScript")
      
      $taskDefinition = $service.NewTask(0)

      if ($UpdateOrRead -eq "Update"){
            $Settings = $taskDefinition.Settings
            $Settings.StartWhenAvailable = $True
            $Settings.DisallowStartIfOnBatteries = $false
            $Settings.ExecutionTimeLimit= "PT5M"

            $triggers = $taskDefinition.Triggers
            $trigger = $triggers.Create(9)
            $trigger.UserId = $env:username

            $InfoTask = $taskDefinition.RegistrationInfo
            $InfoTask.Description=$pid

            $Action = $taskDefinition.Actions.Create(0)
            $Action.Path = "wscript"
            $Action.Arguments= `
                  "`"$PathScriptWork\$($title)_辅助启动.vbs`" --StayInBackgroundWithoutTips"
            
            $UpdateTask = 4
            $null=$rootFolder.RegisterTaskDefinition( `
                  "$env:username\$title\后台开机登录",`
                  $taskDefinition,$UpdateTask,$null,$null, 3)


      }elseif($UpdateOrRead -eq "Read"){
            if (schtasks /query /tn YuphizScript\$env:username\$title\后台开机登录 2>$null){
            $XmlTasks= `
                  [xml]($rootFolder.gettask("$env:username\$title\后台开机登录").xml)
            $DescriptionTasks=$XmlTasks.Task.RegistrationInfo.Description
            return $DescriptionTasks
            }
      }
      
}




#判断后台版是否在运行并更新进程id
function GetOrUpdate-BackgroundProcessId{
      param (
            $GetOrUpdate,
            $IsTips
      )
      $ProcessID=TaskService "Read"
      if ($GetOrUpdate -eq "Get" -and $ProcessID -ne $null){
            if ((get-process -id $ProcessID -erroraction Ignore).ProcessName -eq "powershell"){
                        return $True
                  }else{
                        return $false
                  }
      }elseif ($GetOrUpdate -eq "Get" -and $ProcessID -eq $null){
            return $false
      }elseif ($GetOrUpdate -eq "Update" -and $ProcessID -ne $null){
            if ((get-process -id $ProcessID -erroraction Ignore).ProcessName -eq "powershell"){
                  $null=$Popup.popup("$title 已经运行，不需再启动",0,$null,4096)
                  exit
            }else{
                  TaskService "Update"
                  if ($IsTips -ne "NoTips") {
                        Write-Host "正在运行 ……"
                        $null=$Popup.popup("已启用$title",1,$null,4096)
                  }
            }
      }elseif($GetOrUpdate -eq "Update" -and $ProcessID -eq $null){
            TaskService "Update"
            if ($IsTips -ne "NoTips") {
                  Write-Host "正在运行 ……"
                  $null=$Popup.popup("已启用$title",1,$null,4096)
            }
      }
}





# 运行换色
function changeTheme{
      param (
            $WindowsThemeValue,
            $AppThemeValue
      )
#———— 换Windows模式颜色（开始菜单）
      if ($config.主题颜色选项.切换系统颜色 -eq "开") {
            # $null=$popup.Popup("换系统颜色",1,$null,4096)
            $null = reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v SystemUsesLightTheme /t REG_DWORD /d $WindowsThemeValue /f
      }
#———— 换应用颜色
      if ($config.主题颜色选项.切换应用颜色 -eq "开") {
            # $null=$popup.Popup("换应用颜色",1,$null,4096)
            $null = reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v AppsUseLightTheme /t REG_DWORD /d $AppThemeValue /f
      }

#———— 扩展
      if (($config.扩展选项.启用扩展 -eq "开") -and ((GetOrUpdate-BackgroundProcessId "get") -eq $false)) {
            # $null=$popup.Popup("后台没有开启，运行扩展",1,$null,4096)
            RunExtension
      }
}




# $time1=get-date

function Get-SunRiseSet {
      param (
            $Longitude,
            $Latitude,
            $SunRiseValue=1,
            $dayCount =(New-TimeSpan -Start '2000-01-01 12:00:00').TotalDays,
            $h=[math]::sin(-0.833/180*[math]::PI),
            $ut0=180,
            $ut=0
      )
$LatitudeRadian=$Latitude/180*[math]::PI   #角度转化为弧度

$t= ($dayCount + $ut0 / 360d) / 36525.64 #世纪数
$L = 280.460 + 36000.777 * $t #太阳平均黄径
$G = (357.528 + 35999.050 * $t)/180*[math]::PI #太阳平近点角
$lamda = ($L + 1.915 * [Math]::sin($G) + 0.020 * [Math]::sin(2 * $G))/180*[math]::PI #太阳黄道经度
$epc = (23.4393- 0.0130 * $t)/180*[math]::PI #地球倾角
$sigam = [Math]::asin([Math]::sin($epc) * [Math]::sin($lamda)) #太阳的偏差

$gha = $ut0 - 180 - 1.915 * [Math]::sin($G) - 0.020 * [Math]::sin(2 * $G)+ 2.466 * [Math]::sin(2 * $lamda) - 0.053 * [Math]::sin(4 * $lamda);  # 格林威治时间太阳时间角

$e =([Math]::acos(($h - [Math]::tan($LatitudeRadian) * [Math]::tan($sigam))))*180/[math]::PI  # 修正值e

if ($SunRiseValue -eq 1){
      $ut = $ut0 - $gha - $Longitude - $e
      if ([math]::abs($ut - $ut0) -ge 0.1) {
            $zone=[int][System.TimeZoneInfo]::local.BaseUtcOffset.TotalHours
            $SunRise=($ut / 15 + $zone) + $Config.定位选项.日出时间偏移_时
            Get-SunRiseSet $Longitude $Latitude 1 $dayCount $h $ut $ut
      }else{
            Get-SunRiseSet $Longitude $Latitude 0 $dayCount $h
            return 
      }

}elseif($SunRiseValue -eq 0){
      $ut = $ut0 - $gha - $Longitude + $e
      if ([math]::abs($ut - $ut0) -ge 0.1) {
            Get-SunRiseSet $Longitude $Latitude 0 $dayCount $h $ut $ut
      }else{
            $Sunset=($ut / 15 + $zone) + $Config.定位选项.日落时间偏移_时
            return $SunRise,$SunSet
      }
}
}



# 小时格式化
Function ConvertHourTo($h,$hm){
      $hh=[math]::floor($h) |
            % { if("$_".length -lt 2) {"0"+$_}else{$_}}
      $mm=[math]::floor(($h-$hh)*60) |
            % { if("$_".length -lt 2) {"0"+$_}else{$_}}
      $ss=[math]::floor((($h-$hh)*60-$mm)*60) |
            % { if("$_".length -lt 2) {"0"+$_}else{$_}}
      if ($hm -eq "hm") {
            return "$hh"+":"+"$mm"
      }else{
            return "$hh"+":"+"$mm"+":"+"$ss"
      }
}



#时间转换函数
Function ConvertTimeTo {
      param (
            $TimeDat,
            $What="h"
      )
      $timesplit=$TimeDat.split(":",3)
      if ($What -eq "h"){
            return $timesplit[0]/1+$timesplit[1]/60+$timesplit[2]/3600
      }elseif ($What -eq "m"){
            return $timesplit[0]*60+$timesplit[1]/1+$timesplit[2]/60
      }elseif ($What -eq "s"){
            return $timesplit[0]*3600+$timesplit[1]*60+$timesplit[2]/1
      }
}




#════════════════ 检测任务计划 模块开始 ════════════════
function TimeSchtasks{
      param (
            $EnableOrDisable,
            $SunRise,
            $SunSet,
            $OnlyUpdate  
      )

      $hms_SunRise=ConvertHourTo ($SunRise)
      $hms_SunSet=ConvertHourTo ($SunSet)
# "`n日出时分秒"
# $hms_SunRise
# "`n日落时分秒"
# $hms_SunSet

####主脚本 虽然放到了一起，但是触发器还是延迟启动
$TaskName="\YuphizScript\$env:username\$title\自动主题配色"
if ((! (schtasks /query /tn $TaskName  2>$null)) -and $EnableOrDisable -eq "Enable" ){
      $TriggerClass = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
      $triggerXML = $TriggerClass | New-CimInstance -ClientOnly
      $triggerXML.Enabled = $true
      $triggerXML.Subscription="<QueryList><Query Id='0' Path='System'><Select Path='System'>*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and (Level=4 or Level=0) and (EventID=1)]]</Select></Query></QueryList>"
      $Triggers = @(
            $(New-ScheduledTaskTrigger -daily -at $hms_SunRise),
            $(New-ScheduledTaskTrigger -daily -at $hms_SunSet),
            $(New-ScheduledTaskTrigger -atlogon -user "$env:username"),
            $triggerXML
      )
      $null = Register-ScheduledTask -taskname $TaskName -Action (New-ScheduledTaskAction -Execute "wscript" -Argument """$PathScriptWork\$($title)_辅助启动.vbs"" --RunByTaskWithoutUpdateTime") -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit 00:05 -AllowStartIfOnBatteries)  -Trigger  $Triggers
}else{
      if ($OnlyUpdate -ne "OnlyUpdate") {
      $null = SCHTASKS /change /$EnableOrDisable /tn $TaskName
      }
}


####更新日出日落时间
      $TaskName5="\YuphizScript\$env:username\$title\更新日出日落时间"
      if ((! (schtasks /query /tn $TaskName5  2>$null)) -and $EnableOrDisable -eq "Enable" ){
            $null = schtasks /Create /TN $TaskName5 /TR "wscript '$PathScriptWork\$($title)_辅助启动.vbs' --UpdateSchtasksTime" /SC DAILY /mo 2 /ST 02:00:00 /f
            $null = Set-ScheduledTask -taskname YuphizScript\$env:username\$title\更新日出日落时间 -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit 00:05 -AllowStartIfOnBatteries)
      }else{
            if ($OnlyUpdate -ne "OnlyUpdate") {
                  $null = SCHTASKS /change /$EnableOrDisable /tn $TaskName5
                  }
      }


      if ( $EnableOrDisable -eq "Enable" -and $RunWith -ne "RunByTaskWithoutUpdateTime") {
            $TriggerClass = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
            $triggerXML = $TriggerClass | New-CimInstance -ClientOnly
            $triggerXML.Enabled = $true
            $triggerXML.Subscription="<QueryList><Query Id='0' Path='System'><Select Path='System'>*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and (Level=4 or Level=0) and (EventID=1)]]</Select></Query></QueryList>"
            $Triggers = @(
                  $(New-ScheduledTaskTrigger -daily -at $hms_SunRise),
                  $(New-ScheduledTaskTrigger -daily -at $hms_SunSet),
                  $(New-ScheduledTaskTrigger -atlogon -user "$env:username"),
                  $triggerXML
            )
            $null = Set-ScheduledTask -taskname YuphizScript\$env:username\$title\自动主题配色 -Trigger  $Triggers
      }
}





function ChangeThemeBySchtasks{
      param (
            $SunRise,
            $SunSet,
            $IsTips
      )
      
      $TimeNowh=ConvertTimeTo (Get-Date -Format 'HH:mm:ss')
      
#调试
# $debugtime="19:40"
# $TimeNowh=ConvertTimeTo $debugtime

            
# 判断现在时间 更换变量
      $OffsetSchtask=$Config.其他选项.任务计划偏移时间_时
      if ((($SunRise-$OffsetSchtask) -le $timenowH) -and ($timenowH  -le  ($SunSet-$OffsetSchtask))) {
            $WindowsThemeValue = "1" #系统浅色主题（开始菜单和任务栏）
            $AppThemeValue = "1" #应用浅色主题
      }else{
            $WindowsThemeValue = "0" #系统深色主题（开始菜单和任务栏）
            $AppThemeValue = "0" #应用深色主题
      }
      if ($IsTips -ne "Notips"){
            Write-Host "正在运行 ……"
            $null = $popup.popup("已启用定时版",1,$null,4096)
      }
      changeTheme $WindowsThemeValue $AppThemeValue
}




function BackgroundSchtasks {
      param (
            $EnableOrDisable="Enable"
      )
      $TaskName6="\YuphizScript\$env:username\$title\后台开机登录"
      if ((!(schtasks /query /tn $TaskName6  2>$null)) -and $EnableOrDisable -eq "Enable"){
            $null = Register-ScheduledTask -taskname $TaskName6 -Action (New-ScheduledTaskAction -Execute "wscript" -Argument """$PathScriptWork\$($title)_辅助启动.vbs"" --StayInBackgroundWithoutTips") -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit 00:05 -AllowStartIfOnBatteries)  -Trigger  (New-ScheduledTaskTrigger -atlogon -user "$env:username") 
      }else{
            $null = SCHTASKS /change /$EnableOrDisable /tn $TaskName6
      }
}



function StayInBackground {
      for () {
            Start-sleep -m 700
            $NewSystemTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").SystemUsesLightTheme
            Start-Sleep -m 700
            $NewAppTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme

            if ($NewAppTheme -ne $LastAppTheme -or $NewSystemTheme -ne $LastSystemTheme) {
                  # $null=$popup.popup("NewAppTheme: $NewAppTheme `n`r
# NewSystemTheme: $NewSystemTheme" ,1,$null,4096)
                  if ($NewAppTheme -ne $LastAppTheme ) {
                        $LastAppTheme = $NewAppTheme
                        RunExtension
                        Start-Sleep -m 700
                        $LastSystemTheme = (Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").SystemUsesLightTheme
                  }elseif($NewSystemTheme -ne $LastSystemTheme) {
                        $LastSystemTheme = $NewSystemTheme
                        RunExtension
                        Start-Sleep -m 700
                        $LastAppTheme = (Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme
                  }
                  Start-sleep -m 1000
                  
            }

      }
}

function KillOrRestartBackground {
      param (
            $KillOrRestart
      )
      $ProcessID=TaskService "Read"
      if ($ProcessID -ne $null) {
            if ((get-process -id $ProcessID -erroraction Ignore).ProcessName -eq "powershell"){
                  taskkill /im $ProcessID /f
                  if ($KillOrRestart -eq "AndDisableSchtasks"){
                        BackgroundSchtasks "Disable"
                  }
            }
      }
      if ($KillOrRestart -eq "Restart"){
            BackgroundSchtasks
            GetOrUpdate-BackgroundProcessId "Update"
            StayInBackground
      }
}


function DisableAllSchtasks{
      KillOrRestartBackground "Kill"
      $AllSchTasks=@(
            "\YuphizScript\$env:username\$title\日出浅色",
            "\YuphizScript\$env:username\$title\日落深色",
            "\YuphizScript\$env:username\$title\登录",
            "\YuphizScript\$env:username\$title\唤醒",
            "\YuphizScript\$env:username\$title\更新日出日落时间",
            "\YuphizScript\$env:username\$title\后台开机登录"
            "\YuphizScript\$env:username\$title\壁纸轮播"
      )
      foreach ($Oneof in $AllSchTasks){
            if (schtasks /query /tn $Oneof 2>$null) {
                 $null = schtasks /change /disable /tn $Oneof
            }
      }
      $null = $popup.popup("操作完成",1,$null,4096)
}


function RemoveAllSchtasks{
      $Ask=$popup.popup(
      "防误操作，真的要删除【$title】吗？",
      0, 
      "防误操作，请再确认",
      1+48+256+4096 
  )
  if ($Ask -eq 1) {
      KillOrRestartBackground "Kill"
      $AllSchTasks=@(
            "\YuphizScript\$env:username\$title\日出浅色",
            "\YuphizScript\$env:username\$title\日落深色",
            "\YuphizScript\$env:username\$title\登录",
            "\YuphizScript\$env:username\$title\唤醒",
            "\YuphizScript\$env:username\$title\更新日出日落时间",
            "\YuphizScript\$env:username\$title\后台开机登录",
            "\YuphizScript\$env:username\$title\壁纸轮播",
            "\YuphizScript\$env:username\$title\自动主题配色"
      )
      foreach ($Oneof in $AllSchTasks){
            if (schtasks /query /tn $Oneof 2>$null) {
                 schtasks /delete /tn $Oneof /f
            }
      }

      $service = new-object -com("Schedule.Service")
      $service.connect()
      $rootFolder = $service.Getfolder("\YuphizScript\$env:username")
      $taskDefinition=$service.NewTask(0)
      try{$rootFolder.deleteFolder($title,0)}catch{}
      $null = $popup.popup("操作完成",1,$null,4096)

}else{
      exit
}
}


if ($Config.定位选项.网络自动定位 -eq "开"){ 
      $UrlLocation = "https://api.map.baidu.com/location/ip?ak=HQi0eHpVOLlRuIFlsTZNGlYvqLO56un3&coor=bd09ll"
      try {
            $Location=invoke-restmethod -uri $UrlLocation -UseBasicParsing
      }catch{
         switch ($error[0].FullyQualifiedErrorId) {
             "WebCmdletWebResponseException,Microsoft.PowerShell.Commands.InvokeRestMethodCommand" {
            $null = $popup.popup("    网络错误，定位失败。`
    请检查网络或者手动定位`
    按确定退出……",0,"自动主题配色",4096)
            exit
        }
      }
      }
      $Longitude=$Location.Content.Point.x
      $Latitude=$Location.Content.Point.y
}elseif ($Config.定位选项.网络自动定位 -eq "关" -or $Longitude -eq $null -or $Latitude -eq $null) {
      $Longitude=$Config.定位选项.手动定位经度
      $Latitude=$Config.定位选项.手动定位纬度
}


$SunRiseSet = (Get-SunRiseSet $Longitude $Latitude)
$SunRise = $SunRiseSet[0]
$SunSet = $SunRiseSet[1]

# $SunRise
# $SunSet


switch ($RunWith) {
      "RunByTaskWithoutUpdateTime" {
            ChangeThemeBySchtasks $SunRise $SunSet "Notips"
      }
      "RunChangeByTask" {
            TimeSchtasks "Enable" $SunRise $SunSet
            ChangeThemeBySchtasks $SunRise $SunSet
      }
      "UpdateSchtasksTime" {
            TimeSchtasks "Enable" $SunRise $SunSet "OnlyUpdate"
      }
      "DisableSchtasks" {TimeSchtasks "disable"}
      "RunStayInBackground" {
            BackgroundSchtasks
            GetOrUpdate-BackgroundProcessId "Update"
            StayInBackground
      }
      "RunStayInBackgroundWithoutTips" {
            # BackgroundSchtasks
            GetOrUpdate-BackgroundProcessId "Update" "Notips"
            StayInBackground
      }
      "RestartTheBackground" {
            KillOrRestartBackground "restart"
      }
      "KillTheBackground" {
            KillOrRestartBackground "Kill"
      }
      "KillTheBackgroundAndDisable" {
            KillOrRestartBackground "Kill"
            BackgroundSchtasks "Disable"
      }
      "DisableAllSchtasks" {
            DisableAllSchtasks
      }
      "RemoveAllSchtasks" {
            RemoveAllSchtasks
      }
}