' 声明
' 脚本名称：Yuphiz_自动主题辅助启动器
' 版本号：v1.1.1
' 作者：Yuphiz
' 本脚本是Yuphiz_自动主题脚本的辅助启动
' 凡用此脚本从事法律不允许的事情的，均与本作者无关
' 此脚本许可采用 GPL-3.0-later 协议

Pathcurrentfile = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path

set Shell=CreateObject("WScript.Shell")


select case wscript.arguments.count 
      case 0
            call ScriptLauncherUi()
            ' call IsHaveWallpaperScript(Pathcurrentfile  & "\扩展")
      case 1
            select case wscript.arguments(0)
                  case "--RunByTaskWithoutUpdateTime"
                        call runmain("RunByTaskWithoutUpdateTime",1)
                  case "--UpdateSchtasksTime"
                        call runmain("UpdateSchtasksTime",1)
                  case "--StayInBackgroundWithoutTips"
                        call runmain("RunStayInBackgroundWithoutTips",1)
                  case "--Wallpaper"
                        call RunWallpaperOrNot()
            end select
end select


sub ScriptLauncherUi()
      ask=inputbox( vbcrlf& _
            "※ 定时版：日出日落自动切换配色 (推荐)" &vbcrlf&vbcrlf& _
            "     1     启 用 定 时 版" &vbcrlf&vbcrlf& _
            "     1.2   禁 用 定 时 版"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "※ 后台版：随系统主题颜色启动" &vbcrlf&vbcrlf& _
            "     2     启 用 后 台 版" &vbcrlf&vbcrlf& _
            "     2.1   重 启 (刷新配置用)" &vbcrlf&vbcrlf&vbcrlf& _
            "     2.2   禁 用 后 台 版 (并取消后台开机启动)" &vbcrlf&vbcrlf& _
            "     2.3   仅 停 止 后 台  (下次开机还会启动)" &vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "※   3 进入其他选项 ( 修复 卸载 快捷方式 等)" &vbcrlf&vbcrlf, _
            "自动主题配色   v1.1.1   -- By @Yuphiz",_
            "请输入对应的序号，比如 1")

      select case True
            case Ask=""
                  Wscript.quit
            case Ask="1"
                  call runmain("RunChangeByTask",1)
            case Ask="1.2"
                  call runmain("DisableSchtasks",1)
            case Ask="2"
                  call runmain("RunStayInBackground",01)
            case Ask="2.1"
                  call runmain("RestartTheBackground",01)
            case Ask="2.2"
                  call runmain("KillTheBackgroundAndDisable",1)
            case Ask="2.3"
                  call runmain("KillTheBackground",1)
            case Ask="3"
                  call ScriptLauncherUi2()
                  Wscript.quit
            case else
                  call ScriptLauncherUi()
                  Wscript.quit
      end select
end sub

sub ScriptLauncherUi2()
      ask2=inputbox( vbcrlf& _
            "     1     生 成 设 置 文 件 (第一次启动用)"&vbcrlf&vbcrlf& _
            "     2     升 级 或 修 复 (暂时不可选)" &vbcrlf&vbcrlf&vbcrlf& _
            "     3     禁 用 所 有" &vbcrlf&vbcrlf& _
            "     4     卸 载 所 有" &vbcrlf&vbcrlf&vbcrlf& _
            "     5     创 建 快 捷 方 式" &vbcrlf&vbcrlf& _
            "     6     发 送 快 捷 方 式 到 桌 面" &vbcrlf&vbcrlf, _
            "自动主题配色   v1.1.1   -- By @Yuphiz",_
            "请输入对应的序号，比如 1")    

      select case True
            case Ask2=""
                  Wscript.quit
            case Ask2="1"
                  call runmain("DefaultConfig",1)
            case Ask2="2"
                  ' call runmain("Repair",1)
                  call ScriptLauncherUi2()
                  Wscript.quit
            case Ask2="3"
                  call runmain("DisableAllSchtasks",1)
            case Ask2="4"
                  call runmain("RemoveAllSchtasks",1)
            case Ask2="5"
                  call SendLinkTo(Pathcurrentfile)
            case Ask2="6"
                  DesktopPath = Shell.SpecialFolders("Desktop")
                  call SendLinkTo(DesktopPath)
            case else
                  call ScriptLauncherUi2()
                  Wscript.quit
      end select
end sub


sub runmain(arguments,ishidden)
     Shell.run("powershell -ExecutionPolicy Bypass -File """&Pathcurrentfile &"\自动主题配色.ps1"" "& arguments ),ishidden
end sub



'随机轮播壁纸支持
function IsHaveWallpaperScript(PathFolder)
      set FSO = CreateObject("Scripting.FileSystemObject")
      set Folders=FSO.GetFolder(PathFolder)
      set SubFolders=Folders.SubFolders
      set Files = Folders.Files

      for each Oneof in SubFolders
            FolderArray = split(Oneof,"\")
            FolderName = FolderArray(Ubound(FolderArray))
      ' msgbox FolderName &vbcrlf& Instr(1,FolderName,"__",1)
            if Instr(1,FolderName,"__",1) <> 1 then
                  name = IsHaveWallpaperScript(Oneof)
            ' exit function
            end if
      next
      
      for each Oneof in Files
            if Oneof.Name = "壁纸高级版_Yuphiz.ps1" then
            message= _
                "文件夹：" & Folders &vbcrlf&_
                "文件：" & Oneof.Name &vbcrlf&_
                "文件路径：" & Oneof
                name = Oneof
            end if
            ' msgBox name
            '     exit function
      next
IsHaveWallpaperScript = name
end function



sub RunWallpaperOrNot()
      UserName=CreateObject("WScript.Network").UserName
      Set Shell = CreateObject("Wscript.Shell")
      set FSO = CreateObject("Scripting.FileSystemObject")
      WallpaperScriptPath = IsHaveWallpaperScript(Pathcurrentfile & "\扩展")
      ' msgbox IsHaveWallpaperScriptPath
      if WallpaperScriptPath <> "" then
            ' msgbox 2
            call RandomWallpaper(WallpaperScriptPath,Null,1)
      else
            wscript.quit
      end if
end sub


sub RandomWallpaper(Path,arguments,ishidden)
      Set Shell = CreateObject("Wscript.Shell")
      Shell.run("powershell -ExecutionPolicy Bypass -File """&Path &""" "& arguments ),ishidden
end sub

sub SendLinkTo(Path)
      set LinkObj = Shell.CreateShortcut(Path & "\自动主题配色.lnk")
      LinkObj.TargetPath = "explorer"
      LinkObj.Arguments = Wscript.ScriptFullName
      LinkObj.IconLocation = "%SystemRoot%\System32\shell32.dll,174"
      LinkObj.Save
      Shell.popup "操作完成" &vbcrlf&vbcrlf&_
      Path &vbcrlf&vbcrlf&_
      "创建了快捷方式",2
end sub