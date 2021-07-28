#pragma once
#define _CRT_SECURE_NO_WARNINGS
#define _WIN32_DCOM

#include <windows.h>
#include <stdio.h>
#include <comdef.h>
#include <locale.h>
#include <taskschd.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

#define ERR(func, error)  printf("Error! %s: %d\n", func, error);
#define PingMsgName     L"\\PingDetected.vbs"
#define ChangesMsgName  L"\\SettingChanges.vbs"

enum Mode { Ping_Mode, Update_Mode };   // Режим работы  
int g_taskCount = 0;                    // Счетчик кол-ва задач
LPWSTR g_pingDetectedMsg;               // Строка с путем к скрипту (ping)
LPWSTR g_settingChangesMsg;             // Строка с путем к скрипту (changes)

// Фильтр для Windows Defender и Windows Firewall
wchar_t filterWinSecurity[] = L"<QueryList><Query Id = '0' Path = 'Microsoft-Windows-Windows Defender/Operational'>\
<Select Path = 'Microsoft-Windows-Windows Defender/Operational'>* [System[((EventID &gt;= 5000 and EventID &lt;= 5050))]]\
</Select></Query></QueryList>";

wchar_t filterFirewall[] = L"<QueryList><Query Id = '0' Path = 'Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'>\
<Select Path = 'Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'>* \
[System[Provider[@Name = 'Microsoft-Windows-Windows Firewall With Advanced Security']]]\
</Select></Query></QueryList>";

// Фильтр ping от определенного адреса 
wchar_t filterPing[265] = L"<QueryList><Query Id = '0' Path = 'Security'><Select Path = 'Security'>\
*[System[(EventID = 5152)]] and *[EventData[Data[@Name = 'SourceAddress'] and Data = '%s']]\
</Select></Query></QueryList>";

//Имя задачи
wchar_t wszTaskName[32] = L"Lab3 Trigger Test Task";


