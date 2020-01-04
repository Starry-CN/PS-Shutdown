#版本:0.0.0.1
#2020-1-5 粗创

$script:upShutdown = 0;
$script:dlShutdown = 0;

function Install-Ini
{
  $Source = @"
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    
    
    public class IniFileManager
    {
        [DllImport("KERNEL32.DLL", CharSet = CharSet.Unicode)]
        private static extern uint GetPrivateProfileStringW(
                                                                   string lpAppName,
                                                                   string lpKeyName,
                                                                   string lpDefault,
                                                                   StringBuilder lpReturnedString,
                                                                   uint nSize,
                                                                   string lpFileName);

        [DllImport("KERNEL32.DLL", CharSet = CharSet.Unicode)]
        private static extern uint WritePrivateProfileStringW(
                                                                   string lpAppName,
                                                                   string lpKeyName,
                                                                   string lpString,
                                                                   string lpFileName);

        public static string ReadFromIniFile(string iniFilePath, string appName, string key, string Default)
        {
            StringBuilder sb = new StringBuilder(1024);

            GetPrivateProfileStringW(appName, key, Default, sb, (uint)sb.Capacity, iniFilePath);

            return sb.ToString();
        }

        public static void WriteToIniFile(string iniFilePath, string appName, string key, string value)
        {
            WritePrivateProfileStringW(appName, key, value, iniFilePath);
        }

        public static void UpdateFile(string iniFilePath)
        {
            WritePrivateProfileStringW(null, null, "测试", iniFilePath);
        } 
}    
"@;
  #$ref=@("System.IO", "System.Text");
  #Add-Type -MemberDefinition $Source -Name "IniFile" -using $ref -PassThru; 
  Add-Type -TypeDefinition $Source;  
};

function Read-Ini($IniKey)
{
  $Section = "关机";
  $FilePath = "$PSScriptRoot\Site.ini";
  $Val = "-999";
  [IniFileManager]::ReadFromIniFile($FilePath, $Section, $IniKey, $Val);
};

function UpdateFile-Ini
{
  $FilePath = "$PSScriptRoot\Site.ini";
  $Null = [IniFileManager]::UpdateFile($FilePath);  
};

function Write-Ini
{
  param ([string]$IniKey, [string]$Value);
  $Section = "关机";
  $FilePath = "$PSScriptRoot\Site.ini";
  $Null = [IniFileManager]::WriteToIniFile($FilePath, $Section, $IniKey, $Value); 
};

function Get-NetworkName
{ 
  $category = New-Object -TypeName Diagnostics.PerformanceCounterCategory -Property @{CategoryName = "Network Interface"};
  $names = $category.GetInstanceNames();
  Write-Host "";
  Write-Host "本机网卡：";
  Write-Host "DeviceID Name";
  Write-Host "-------- ----";
  for ($i=0;$i -lt $names.length;$i++)
  {
    $name = $names[$i];
    if ($name -eq "MS TCP Loopback interface")
    {
      continue;
    }else{
      Write-Host "$i        $name"; 
    };
    $name;
  };
};

function Get-NetworkSpeed($NetworkName)
{
  $dlCounter = New-Object -TypeName Diagnostics.PerformanceCounter -Property @{CategoryName = "Network Interface"; CounterName = "Bytes Received/sec"; InstanceName = $NetworkName};
  $ulCounter = New-Object -TypeName Diagnostics.PerformanceCounter -Property @{CategoryName = "Network Interface"; CounterName = "Bytes Sent/sec"; InstanceName = $NetworkName};
  $dlValueOld = $dlCounter.NextSample().RawValue;
  $ulValueOld = $ulCounter.NextSample().RawValue;
  $i = 0;
  $d1 = Get-Date;
  $d2 = Get-Date;
  $Minutes3 = 0;
  while ($true){
    $dlValue = $dlCounter.NextSample().RawValue;
    $ulValue = $ulCounter.NextSample().RawValue; 
    $dlSpeed = ($dlValue - $dlValueOld)/1024;
    $ulSpeed = ($ulValue - $ulValueOld)/1024;  
    $dlValueOld = $dlCounter.NextSample().RawValue;
    $ulValueOld = $ulCounter.NextSample().RawValue;       
    $DownloadSpeed =  ("{0:N2}" -f ($dlSpeed));
    $UploadSpeed =  ("{0:N2}" -f ($ulSpeed));
    if ($dlShutdown -gt 0){
      if ($dlSpeed -lt $dlShutdown){if ($i -eq 0){$d1 = Get-Date;};$i = 1;};
      if ($dlSpeed -gt $dlShutdown){$i = 0;};
    }else{$d1 = $d2};
    if ($upShutdown -gt 0){
      if ($ulSpeed -lt $upShutdown){if ($i -eq 0){$d2 = Get-Date;};$i = 1;};
      if ($dlSpeed -gt $upShutdown){$i = 0;};
    }else{$d2 = $d1};
    if ($i -eq 1){$Minutes1 = (New-TimeSpan -Start $d1 -End (Get-Date)).TotalMinutes}; 
    if ($i -eq 1){$Minutes2 = (New-TimeSpan -Start $d2 -End (Get-Date)).TotalMinutes};
    if ($Minutes1 -ge 5){$i = 88};
    if ($Minutes2 -ge 5){$i = 88};
    if ($i -eq 88){cls;Write-Host "大约2分钟后关机";break;};
    if ($Minutes1 -ge $Minutes2){$Minutes3 = ("{0:N2}" -f $Minutes1)}; 
    if ($Minutes2 -ge $Minutes1){$Minutes3 = ("{0:N2}" -f $Minutes2)};
    cls;
    Write-Host ""; 
    #Write-Host $d1 $d2 $i; 
    if ($Minutes3 -gt 0){Write-Host "已经 $Minutes3 分钟，网速低于阈值"};
    Write-Host "";
    Write-Host "";
    Write-Host "此窗口正在监测网速中，请勿关闭！！！"; 
    Write-Host $NetworkName：; 
    Write-Host "当前网卡的下载速度：$DownloadSpeed KB/s";
    Write-Host "当前网卡的上传速度：$UploadSpeed KB/s";
    sleep 1;
  };
  shutdown -s -t 160;
};

function Main
{
  param ([string]$Option="");
  Install-Ini;
  $names = Get-NetworkName; 
  $names = [String[]]$names;
  $up = Read-Ini("UploadSpeed");
  Write-Host "";  
  if ($up -eq "-999")
  {
    Write-Host "请设置关机下载速度";
    $dlShutdown = Read-Host "五分钟内小于此速度将关机";
    Write-Host "请设置关机上传速度";
    $upShutdown = Read-Host "五分钟内小于此速度将关机";
    Write-Ini -IniKey "DownloadSpeed" -Value $dlShutdown;
    Write-Ini -IniKey "UploadSpeed" -Value $upShutdown;
    UpdateFile-Ini;
  }else{
    $dlShutdown = Read-Ini("DownloadSpeed");
    $upShutdown = Read-Ini("UploadSpeed");
  };

  if ($names.length -gt 1)
  {
    $i = Read-Host "请选择检测网卡";
  }else{
    $i = 0
  };
  if ($i -ge 0)
  {
    if ($i -lt $names.length)
    {
      $name = $names[$i];
    }else{
      menur1;
    };
  }else{
    menur1;
  };
  
  Get-NetworkSpeed($name);
};

function shutdown3
{
  param ([string]$Option);
  $d1 = Get-Date;
  $s1 = $Option -split ":";
  $s2 = [string]$s1[0] -split "\.";
  $Day1 = $s2[0];
  $Hour1 = $s2[1];
  $Minute1 = $s1[1];
try{
    $Seconds1 =  (New-TimeSpan -Days $Day1 -Hours $Hour1 -Minutes $Minute1).TotalSeconds;
    if ($Seconds1 -le 0){ menu3 -Option 1; };
    if ($Seconds1 -ge 315360000){ menu3 -Option 1; };
}
catch{
    Write-Host "错误:"$_;
    menu3 -Option 1;

};
  cls;
  Write-Host '关机程序启动中';
  shutdown -s -t $Seconds1;
  Write-Host '感谢您，使用本程序！';
  sleep 6;
};

function shutdown2
{
  param ([string]$Option);
  $d1 = Get-Date;
  $s1 = $Option -split ":";
  $s2 = [string]$s1[0] -split "\.";
  $Day1 = $s2[0];
  $Hour1 = $s2[1];
  $Minute1 = $s1[1];
try{
    $d2 = (Get-Date -Day $Day1 -Hour $Hour1 -Minute $Minute1);
    if ($d2.Day -eq $Day1){
      $Seconds1 =  (New-TimeSpan -End $d2).TotalSeconds;
      if ($Seconds1 -le 0)
      {
        $m1 = $d1.Month;
        if ($m1 -lt 12)
        {
          $m1 = $m1 + 1;
        }else{
          $m1 = 1;
        };
        $d2 = (Get-Date -Month $m1 -Day $Day1 -Hour $Hour1 -Minute $Minute1);
        if ($d2.Day -eq $Day1){
          $Seconds1 =  (New-TimeSpan -End $d2).TotalSeconds;
        }else{ menu2 -Option 1; };
        
      };
    }else{ menu2 -Option 1; };
    if ($Seconds1 -le 0){ menu2 -Option 1; };
    if ($Seconds1 -ge 315360000){ menu2 -Option 1; };
}
catch{
    Write-Host "错误:"$_;
    menu2 -Option 1;

};
  cls;
  Write-Host '关机程序启动中';
  shutdown -s -t $Seconds1;
  Write-Host '感谢您，使用本程序！';
  sleep 6;
};

function menu3
{
  param ([int]$Option=0);
  cls;
  Write-Host ""
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen 
  Write-Host "2.经过某天某小时某分钟关机"
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "请输入以下格式`"00.00:00`"，本程序会启动关机程序                  " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "例如`"2.22:6`"，代表2天22小时6分钟后将会关机                      " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "2.输入`"2`"，将返回主菜单                                         " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen
  Write-Host ""
  if ($Option -eq 0){Write-Host "请输入您的选择：";};
  if ($Option -eq 1){Write-Host "请重新输入您的选择：";};
  $str1 = Read-Host ;
  $Regular1 = "\d{1,2}\.\d{1,2}:\d{1,2}";
  $i = $str1 -replace '`"', "";
  $i = $str1 -replace '：', ":";
  $i = $str1 -replace ';', ":";
  $i = $str1 -replace '/', ":";
  if ($i -match $Regular1) 
  {
    $e1 = shutdown3 -Option $i;
  }
  elseif ($i -eq "2")
  {
    menu1;
  };
};

function menu2
{
  param ([int]$Option=0);
  cls;
  Write-Host ""
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen 
  Write-Host "1.在某日某时某分关机"
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "请输入以下格式`"00.00:00`"，本程序会启动关机程序                  " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "例如`"29.22:6`"，代表本月29日22时6分将会关机                      " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "如果本月29日已经过，那么将会在下月29日自动关机                  " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "    如果下月没有29日，那么程序会自动返回此菜单                  " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "2.输入`"2`"，将返回主菜单                                         " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen
  Write-Host ""
  if ($Option -eq 0){Write-Host "请输入您的选择：";};
  if ($Option -eq 1){Write-Host "请重新输入您的选择：";};
  $str1 = Read-Host ;
  $Regular1 = "\d{1,2}\.\d{1,2}:\d{1,2}";
  $i = $str1 -replace '`"', "";
  $i = $str1 -replace '：', ":";
  $i = $str1 -replace ';', ":";
  $i = $str1 -replace '/', ":";
  if ($i -match $Regular1) 
  {
    $e1 = shutdown2 -Option $i;
  }
  elseif ($i -eq "2")
  {
    menu1;
  };
};

function menu1
{
  cls;
  Write-Host ""
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen 
  Write-Host "本软件的功能是计划关机"
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "目前已实现的功能有："
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "1.在某日某时某分关机（如本月此日已过，那就下月此日）            " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "2.经过某天某小时某分钟关机（例：在5分钟后关机）                 " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "3.在网速低于某点长达5分钟时，自动关机                           " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "   （3.1.注意：以网速为关机条件，此窗口必须打开，不能关闭）     " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "4.修改参数                                                      " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "   （4.1.本软件在使用过程会自动记录相关参数，以便下次自动执行） " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "                                Copyright (c) 2020, Starry-CN.  "
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen
  Write-Host ""
  Write-Host "请输入您的选择："
  $i = Read-Host ;
  $Regular1 = "\d";
  if ($i -match $Regular1) 
  {
    if ($i -gt 4)
    {
      menur1;
    }
    elseif ($i -lt 1)
    {
      menur1;
    }
    else
    {
      switch ($i)
      {
        1 {  menu2; }
        2 {  menu3; }
        3 {  Main;  }
        4 {"It is four."}
      };
    };
  };
};

function menu1
{
  cls;
  Write-Host ""
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen 
  Write-Host "本软件的功能是计划关机"
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "目前已实现的功能有："
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "1.在某日某时某分关机（如本月此日已过，那就下月此日）            " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "2.经过某天某小时某分钟关机（例：在5分钟后关机）                 " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "3.在网速低于某点长达5分钟时，自动关机                           " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "   （3.1.注意：以网速为关机条件，此窗口必须打开，不能关闭）     " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
#  Write-Host "4.修改参数                                                      " -ForegroundColor DarkGreen -BackgroundColor Black
#  Write-Host "   （4.1.本软件在使用过程会自动记录相关参数，以便下次自动执行） " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "                                                                " -ForegroundColor DarkGreen -BackgroundColor Black
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkRed
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkBlue
  Write-Host "                                Copyright (c) 2020, Starry-CN.  "
  Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGreen
  Write-Host ""
  Write-Host "请输入您的选择："
  $i = Read-Host ;
  $Regular1 = "\d";
  if ($i -match $Regular1) 
  {
    if ($i -gt 4)
    {
      menur1;
    }
    elseif ($i -lt 1)
    {
      menur1;
    }
    else
    {
      switch ($i)
      {
        1 {  menu2; }
        2 {  menu3; }
        3 {  Main;  }
        4 {"It is four."}
      };
    };
  };
};


shutdown -a
menu1;

