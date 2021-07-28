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

enum Mode { Ping_Mode, Update_Mode };   // ����� ������  
int g_taskCount = 0;                    // ������� ���-�� �����
LPWSTR g_pingDetectedMsg;               // ������ � ����� � ������� (ping)
LPWSTR g_settingChangesMsg;             // ������ � ����� � ������� (changes)

// ������ ��� Windows Defender � Windows Firewall
wchar_t filterWinSecurity[] = L"<QueryList><Query Id = '0' Path = 'Microsoft-Windows-Windows Defender/Operational'>\
<Select Path = 'Microsoft-Windows-Windows Defender/Operational'>* [System[((EventID &gt;= 5000 and EventID &lt;= 5050))]]\
</Select></Query></QueryList>";

wchar_t filterFirewall[] = L"<QueryList><Query Id = '0' Path = 'Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'>\
<Select Path = 'Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'>* \
[System[Provider[@Name = 'Microsoft-Windows-Windows Firewall With Advanced Security']]]\
</Select></Query></QueryList>";

// ������ ping �� ������������� ������ 
wchar_t filterPing[265] = L"<QueryList><Query Id = '0' Path = 'Security'><Select Path = 'Security'>\
*[System[(EventID = 5152)]] and *[EventData[Data[@Name = 'SourceAddress'] and Data = '%s']]\
</Select></Query></QueryList>";

//��� ������
wchar_t wszTaskName[32] = L"Lab3 Trigger Test Task";


