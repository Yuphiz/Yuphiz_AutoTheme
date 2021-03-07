<#
.声明
脚本：UWP启动页颜色_扩展版
版本：v0.2
作者：Yuphiz
本脚本可以帮助 Windows10 自动切换UWP启动页的颜色
凡用此脚本从事法律不允许的事情的，均与本作者无关
此脚本许可采用 GPL-3.0-later 协议
#>

param ($Color,$ReStore,$TurnOnStayBackstage=0,$IsTips="NoTips")


#脚本所在路径，如果为空则选择当前工作路径
$PathScriptWork = $PSScriptRoot;if ($PathScriptWork -eq "") {$PathScriptWork=(get-location).path}
$title="UWP启动页颜色"
$titleEN="SplashScreen"
$Popup=new-object -Comobject Wscript.Shell
write-host "正在运行 UWP 启动页颜色 ……"

$FileConfig = "$PathScriptWork\UWP启动页颜色_Yuphiz.Json"
if (!(Test-Path $FileConfig)){
      @'
{
      "颜色来源":"自定义"
,     "自定义来源选项":{
            "浅色主题颜色":"#F8F8FF"
,           "深色主题颜色":"#303033"
      }
,     "排除应用":[
      ]
,     "备份":"开_暂时不可选"
}
'@ | set-content $FileConfig
}
try {
      $Config=get-content $FileConfig | ConvertFrom-Json
}catch{
      $null = $popup.popup("      $FileConfig `n`r
$($error[0])",0,"配置文件出错",16);exit
}

#════════════════════ 设置 模块开始 ═════════════════════
# 颜色来源，"Custom"自定义，"System"来自系统主题色，"Random"颜色随机（暂时无效）
$ColorFrom = $Config.颜色来源

#浅色主题颜色
$ColorOfLight = $Config.自定义来源选项.浅色主题颜色
#深色主题颜色
$ColorOfDark = $Config.自定义来源选项.深色主题颜色

#开启备份，建议开启
$TurnOnBackup = $Config.备份


#排除应用，注意软件全名、引号、逗号
$Exclude = $Config.排除应用
#————————————— 设置 模块结束 ——————————————



#════════════════════ 恢复 模块开始 ═════════════════════
if ($ReStore -eq "ReStoreFromUpdate" -or $ReStore -eq "ReStoreFromOriginal") {
   
#════════ 恢复颜色操作 开始
      function RunReStoreColor{
            param(
                  $PackageFamilyName,
                  $nameid,
                  $color
      )
            $null=reg add "$regROOT\$PackageFamilyName\SplashScreen\$nameid" /v BackgroundColor /t REG_SZ /d $color /f

            $key=$regROOT+"\$PackageFamilyName\SplashScreen\$nameid"
            $BackgroundColor=$color

            $NameIDHash=@{}
      return $NameIDHash  | Select-Object -Property  @{label="名称";expression={$nameid}},@{label="注册表路径";expression={$key}},@{label="启动界面颜色";expression={$BackgroundColor}}
      }
#—————— 恢复颜色操作 结束

#═════════ 获取恢复颜色 开始
      function ReStoreColor{
            param(
                  $WhereFileRestore
            )
            $FileJsonToRestore = switch ($WhereFileRestore) {
                  "ReStoreFromUpdate" {
                        "$PathScriptWork\颜色备份\UWPSplashScreenColor_默认备份_更新_"+`
                        $env:userdomain+"_"+$env:username+".json" 
                   }
                  "ReStoreFromOriginal" {
                        "$PathScriptWork\颜色备份\UWPSplashScreenColor_默认备份_原始_勿删_"+`
                        $env:userdomain+"_"+$env:username+".json"
                  }
            }

            $RestoreSplashScreen=get-content $FileJsonToRestore |
                  ConvertFrom-Json
            if($Popup.popup("      恢 复 来 源 是……`n`
      $FileJsonToRestore`n`
      确 定 要 恢 复 吗？",0,"请确认……",1+32+256+4096) -eq 2) {
                  $Popup.popup("你取消了恢复操作")
                  exit
            }

$regROOT="HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData"
            $result=@()
            foreach ($i in $RestoreSplashScreen.nameID){
                  $namei=$RestoreSplashScreen.contents.$i.name
                  $colori=$RestoreSplashScreen.contents.$i.color
                  if (test-path ( `
                        "registry::"+$RegROOT+"\$namei\SplashScreen\$i")){
                        $result+=$(RunReStoreColor $namei $i $colori)
                  }
            }

            $result | Out-GridView -title "还原结果如下，可以直接复制结果" -wait
      }
#——————获取恢复颜色 结束

      ReStoreColor $ReStore
      exit
}
#————————————— 恢复 模块结束 ——————————————




#════════════════════ 获取颜色 开始 ════════════════════
Function getColor($AppTheme){
      $Color = switch ($ColorFrom){
            "壁纸"{
                  "transparent"
                  # "#"+("{0:x}" -f ((Get-ItemProperty -path "registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM").ColorizationColor)).SubString(2,("{0:x}" -f ((Get-ItemProperty -path "registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM").ColorizationColor)).Length-2)
               }
            "自定义"{
                  switch ($AppTheme) {
                        1 { $ColorOfLight }
                        0 { $ColorOfDark }
                  }
               }
            "系统"{
                  "transparent"
               }
            "随机"{}
      }
Return $Color
}
#————————————— 获取颜色 结束 —————————————

$AppTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme
$Color=(getColor $AppTheme)
#════════════════════ 更换颜色 开始 ════════════════════
function RunChangeColor{
      param(
            $PackageFamilyName,
            $id
      )
$null = reg add "$regROOT\$PackageFamilyName\SplashScreen\$PackageFamilyName!$id" /v BackgroundColor /t REG_SZ /d $Color /f
}
#————————————— 更换颜色 结束 —————————————



#════════════════════ 遍历UWP 开始 ════════════════════
Function CheckColor{
#$containsUWP+=$Exclude
$JsonSplashScreen=@()
$JsonNameID=@()
$UWPNameID=@()
$NameArray=@()
$regROOT="HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData"
#$PathRegROOT="HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData"

$AllAppx = get-appxpackage -PackageTypeFilter main | ?{$_.name -notmatch "Microsoft.LanguageExperiencePack"}
foreach ($i in $AllAppx){
      $NameUWP=$i.PackageFamilyName
      $PathUWP=$i.InstallLocation
      $XML =[xml](Get-Content "$PathUWP\AppxManifest.xml")
      [string[]]$ApplicationID=$xml.Package.Applications.Application.id
      if (($NameUWP -ne $null) -and ($ApplicationID -ne $null) `
          -and ($Exclude -notcontains $i.name) -and ($NameUWP -notlike "Microsoft.*Extension*") -and ($NameUWP -notlike "Microsoft.LanguageExperiencePack*")){
#            $NameUWP
#            $ApplicationID
            for ($i=0;$i -lt $ApplicationID.count; $i++){
                  if (test-path ("registry::"+$regROOT+"\$NameUWP\SplashScreen\$NameUWP!$($ApplicationID[$i])")){
                        $BackgroundColor= ( Get-ItemProperty -path ("registry::"+$regROOT+"\$NameUWP\SplashScreen\$NameUWP!$($ApplicationID[$i])")).BackgroundColor
                        if ($NameUWP -ne $null -and $BackgroundColor -ne $null) {
                              $JsonSplashScreen+=@"
{"$NameUWP!$($ApplicationID[$i])":{"name":"$NameUWP","id":"$($ApplicationID[$i])","color":"$BackgroundColor"}},
"@
$JsonNameID+=@"
"$NameUWP!$($ApplicationID[$i])",
"@
                              $UWPNameID+="$NameUWP!$($ApplicationID[$i])"
                              RunChangeColor $NameUWP $ApplicationID[$i]
                        }
                  }
#                  $count=$count+1
            }
      }
}

#$count;$count=0
$JsonSplashScreen="$JsonSplashScreen".Trim(" .-`t`n`r,")
$JsonNameID="$JsonNameID".Trim(" .-`t`n`r,")
$JsonSplashScreen="{""contents"":[$JsonSplashScreen],""NameID"":[$JsonNameID]}"

if (!(test-path "$PathScriptWork\颜色备份")) {
      new-item "$PathScriptWork\颜色备份"  -itemtype "directory"
}
$FileJsonBakOrig="$PathScriptWork\颜色备份\UWPSplashScreenColor_默认备份_原始_勿删_"+$env:userdomain+"_"+$env:username+".json"
$FileJsonBak="$PathScriptWork\颜色备份\UWPSplashScreenColor_默认备份_更新_"+$env:userdomain+"_"+$env:username+".json"

if (!(test-path $FileJsonBakOrig)){
      $JsonSplashScreen >$FileJsonBakOrig
}

if (!(test-path $FileJsonBak)){
      $JsonSplashScreen >$FileJsonBak
}else{
      $ScriptSplashScreen=$JsonSplashScreen | ConvertFrom-Json

      $BakSplashScreen=get-content $FileJsonBak | ConvertFrom-Json

      $isexist=0
      foreach ($i in $UWPNameID){
            if ($BakSplashScreen.NameID -notcontains $i) {
                  $BakSplashScreen.NameID+=,$i
                  $jsoni=@{}
                  $jsoni.name=$ScriptSplashScreen.contents.$i.name
                  $jsoni.id=$ScriptSplashScreen.contents.$i.id
                  $jsoni.color=$ScriptSplashScreen.contents.$i.color
                  $jsonis=@{}
                  $jsonis.$i=$jsoni
                  $BakSplashScreen.contents+=,$jsonis
                  $isexist++
            }
      }

      if ($isexist -ne 0 ) {
            $BakSplashScreen | ConvertTo-Json -Depth 10 | set-content      $FileJsonBak
            "有不同"
            $isexist
      } else{
            "都相同"
            #$isexist
      }
}
}
#————————————— 遍历UWP 结束 —————————————

CheckColor

#════════════════════ 弹窗调试 开始 ════════════════════
function debugpopup {
      $debugAppTheme=(Get-ItemProperty -path "registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize").AppsUseLightTheme
      if ($debugAppTheme -eq 1) {
            $deapptheme="浅色"
      }else{
            $deapptheme="深色"
      }
$null=$Popup.popup("主题是 "+$deapptheme+"`n颜色是 "+$color+"`n颜色来自 "+$colorfrom+"`n白天颜色 "+$ColorOfLight+"`n夜晚颜色 "+$ColorOfDark,3,$null,4096)
}
#————————————— 弹窗调试 结束 —————————————
# debugpopup